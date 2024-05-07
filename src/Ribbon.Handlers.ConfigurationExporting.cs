using MSExcel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System.Text.Json.Nodes;
using KAT.Camelot.RBLe.Core.Calculations;
using KAT.Camelot.Domain.Extensions;
using System.Globalization;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

enum LookupConfigurationType
{
	DataTables,
	FrameworkTables,
	RateTables,
	GlobalTables
}

public partial class Ribbon
{
	public void ConfigurationExporting_ExportWorkbook( IRibbonControl _ )
	{
		skipProcessSaveHistory = true;

		if ( WorkbookState.IsGlobalTablesFile )
		{
			ExportGlobalTables();
		}
		else if ( WorkbookState.IsRTCFile )
		{
			MessageBox.Show( "// TODO: Export RTCFile" );
		}
		else if ( WorkbookState.IsSpecSheetFile )
		{
			MessageBox.Show( "// TODO: Export SpecSheetFile" );
		}
	}

	public void ConfigurationExporting_ProcessGlobalTables( IRibbonControl _ )
	{
		var existing = application.GetWorkbook( Constants.FileNames.GlobalTables );
		var downloadName = existing == null
			? Path.Combine( AddIn.ResourcesPath, Constants.FileNames.GlobalTables )
			: null; // Don't download...

		ExportGlobalTables( downloadName );
	}

	public void ConfigurationExporting_ExportSheet( IRibbonControl _ )
	{
		if ( WorkbookState.SheetState.IsGlobalTableSheet )
		{
			ExportGlobalTables( currentSheet: !WorkbookState.IsGlobalTablesFile );
		}
		else if ( WorkbookState.SheetState.IsUserAccessSheet )
		{
			MessageBox.Show( "// TODO: Process UserAccessSheet" );
		}
	}

	private void ExportGlobalTables( string? downloadName = null, bool currentSheet = false )
	{
		var owner = new NativeWindow();
		owner.AssignHandle( new IntPtr( application.Hwnd ) );

		RunRibbonTask( async () =>
		{
			using var processGlobalTables = new ProcessGlobalTables( 
				currentSheet,
				AddIn.Settings.DataServices,
				GetWindowConfiguration( nameof( ProcessGlobalTables ) )				
			);

			var info = processGlobalTables.GetInfo( 
				AddIn.Settings.KatUserName, 
				await AddIn.Settings.GetClearPasswordAsync(),
				owner
			);

			if ( info == null ) return;

			ExcelAsyncUtil.QueueAsMacro( () => application.Cursor = MSExcel.XlMousePointer.xlWait );

			SaveWindowConfiguration( nameof( ProcessGlobalTables ), info.WindowConfiguration );
			
			await UpdateAddInCredentialsAsync( info.UserName, info.Password );

			await DownloadLatestCalcEngineAsync( downloadName );

			SetStatusBar( "Processing Global Tables..." );

			ExcelAsyncUtil.QueueAsMacro( () =>
			{
				application.Cursor = MSExcel.XlMousePointer.xlWait;
				var existing = currentSheet ? application.ActiveWorkbook : application.GetWorkbook( Constants.FileNames.GlobalTables )!;

				var sheets = existing.Worksheets.Cast<MSExcel.Worksheet>().ToDictionary( x => x.Name );
				var resourceTable = sheets.TryGetValue( Constants.SpecSheet.TabNames.Localization, out var l )
					? l.RangeOrNull( Constants.SpecSheet.RangeNames.ResourceTable )
					: null;

				var globalSpecifications = new JsonObject();

				static JsonObject getGlobalTables( MSExcel.Worksheet worksheet )
				{
					var tables = GetGlobalLookupTables( worksheet ).ToJsonArray();
					var version = worksheet.RangeOrNull<string>( Constants.SpecSheet.RangeNames.SheetVersion );

					return new JsonObject
					{
						{ "version", version },
						{ "tables", tables }
					};
				}

				if ( currentSheet )
				{
					var worksheet = application.ActiveWorksheet();
					
					var propertyName = worksheet.RangeOrNull<string>( Constants.SpecSheet.RangeNames.SheetType ) == Constants.SpecSheet.SheetTypes.GlobalLookupTables
						? "dataTables"
						: "rateTable";

					globalSpecifications[ propertyName ] = getGlobalTables( worksheet );
				}
				else
				{
					globalSpecifications[ "dataTables" ] = getGlobalTables( sheets[ Constants.SpecSheet.TabNames.DataLookupTables ] );
					globalSpecifications[ "rateTables" ] = getGlobalTables( sheets[ Constants.SpecSheet.TabNames.RateTables ] );
				}

				RunRibbonTask( async () => {
					var validations = await apiService.UpdateGlobalTablesAsync( info.ClientName, info.Targets, globalSpecifications, info.UserName, info.Password );

					if ( validations != null )
					{
						ShowValidations( validations );
						return;
					}

					ExcelAsyncUtil.QueueAsMacro( () =>
					{
						MessageBox.Show( owner, $"Successfully updated Global Tables on following environments.{Environment.NewLine + Environment.NewLine + string.Join( ", ", info.Targets )}", "Update Global Tables" );
						
						if ( downloadName != null )
						{
							existing.Close( SaveChanges: false );
						}
					} );
				}, "ExportGlobalTables (Processing)" );
			} );
		} );
	}

	private static IEnumerable<JsonObject> GetGlobalLookupTables( MSExcel.Worksheet worksheet )
	{
		var configurationType = worksheet.RangeOrNull<string>( Constants.SpecSheet.RangeNames.SheetType ) switch
		{
			Constants.SpecSheet.SheetTypes.GlobalLookupTables => LookupConfigurationType.GlobalTables,
			Constants.SpecSheet.SheetTypes.GlobalRateTables or Constants.SpecSheet.SheetTypes.ClientRateTables => LookupConfigurationType.RateTables,
			_ => LookupConfigurationType.DataTables // TODO: Is this right?
		};

		var headerOffset = configurationType == LookupConfigurationType.DataTables ? 2 : 1;

		var firstColumn = 
			worksheet.Range[ worksheet.RangeOrNull<string>( Constants.SpecSheet.RangeNames.TableStartAddress ) ]
				.GetReference()
				.Offset( headerOffset, 0 );

		while ( !string.IsNullOrEmpty( firstColumn.GetText() ) )
		{
			var tableName = firstColumn.Offset( -headerOffset, 1 ).GetText()!;
			var lastColumn = firstColumn.End( DirectionType.ToRight );
			var tableInclude = firstColumn.Offset( -1, 1 ).GetText() ?? "Y";

			if ( configurationType != LookupConfigurationType.DataTables || tableInclude.StartsWith( "Y" ) )
			{
				var lastRow = firstColumn.End( DirectionType.Down );
				var data =
					firstColumn
						.Extend( lastColumn )
						.Extend( lastRow )
						.GetValueArray();

				if ( tableInclude.Contains( "/customize" ) )
				{
					throw new NotImplementedException( $"No support for /customize. {tableName} used this flag.  See if it is needed." );
				}

				yield return new JsonObject
				{
					{ "name", tableName },
					{ "columns", new JsonArray().AddItems( data.Rows.First() ) },
					{ "rows", data.Rows.Skip( 1 ).Select( r => new JsonArray().AddItems( r.Select( GetExportValue ), includeNulls: true ) ).ToJsonArray() }
				};
			}

			firstColumn = lastColumn.End( DirectionType.ToRight, ignoreEmpty: true );
		}
	}

	private static string? GetExportValue( object value )
	{
		if ( value == ExcelEmpty.Value ) return null;

		var d = value as double?;
		if ( d != null )
		{
			// .NET Core changed in the underlying implementation of the Double.ToString() method. 
			// In .NET Core, the method has been updated to produce a round-trippable result by default, 
			// which means it will always return a string that, when parsed, will produce the original number.
			// To reproduce what we had in .NET Framework it is suggested to use the G15 format.  
			// The "G" format specifier stands for "general", and it formats the number in the most compact, human-readable form.
			// In .NET Framework, the default precision for double.ToString() without any format specifier is up to 15 digits, 
			// which can be either to the left or right of the decimal point. This means it can include up to 15 significant digits, 
			// and the remaining digits are replaced with zeros.
			return d.Value.ToString( "G15", CultureInfo.InvariantCulture );

			// I was going to count the number of decimal places and use that + 6, but it seems to be a bit overkill.
			// var decimalValues = ( (int)Math.Floor( Math.Abs( d.Value ) ) ).ToString().Length;
			// return d.Value.ToString( $"G{decimalValues + 6}", CultureInfo.InvariantCulture );
		}

		var dt = value as DateTime?;
		if ( dt != null ) return dt.Value.ToString( "yyyy-MM-dd" );

		var s = (string)value;

		// Excel does all caps for these.
		if ( s == "TRUE" ) s = "true";
		else if ( s == "FALSE" ) s = "false";

		return s;
	}
}