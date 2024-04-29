using MSExcel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System.Text.Json.Nodes;
using KAT.Camelot.RBLe.Core.Calculations;
using KAT.Camelot.Domain.Extensions;

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
		skipHistoryUpdateOnMoveSpecFromDownloads = true;

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
		RunRibbonTask( async () =>
		{
			using var processGlobalTables = new ProcessGlobalTables( 
				currentSheet,
				AddIn.Settings.DataServices,
				GetWindowConfiguration( nameof( ProcessGlobalTables ) ) 
			);

			var info = processGlobalTables.GetInfo( 
				AddIn.Settings.KatUserName, 
				await AddIn.Settings.GetClearPasswordAsync() 
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

				var localization = GetGlobalResourceStrings( resourceTable );

				var globalSpecifications = new JsonObject
				{
					{ "localization", localization }
				};

				if ( currentSheet )
				{
					var worksheet = application.ActiveWorksheet();
					
					var propertyName = worksheet.RangeOrNull<string>( Constants.SpecSheet.RangeNames.SheetType ) == Constants.SpecSheet.SheetTypes.GlobalLookupTables
						? "lookupTables"
						: "rateTable";

					globalSpecifications[ propertyName ] = GetLookupTables( worksheet ).ToJsonArray();
				}
				else
				{
					globalSpecifications[ "lookupTables" ] = GetLookupTables( sheets[ Constants.SpecSheet.TabNames.DataLookupTables ] ).ToJsonArray();
					globalSpecifications[ "rateTables" ] = GetLookupTables( sheets[ Constants.SpecSheet.TabNames.RateTables ] ).ToJsonArray();
				}

				RunRibbonTask( async () => {
					await Task.Delay( 1000 );

					ExcelAsyncUtil.QueueAsMacro( () =>
					{
						MessageBox.Show( "Done Processing" );
						if ( downloadName != null )
						{
							existing.Close( SaveChanges: false );
						}
					} );
				}, "ExportGlobalTables (Processing)" );
			} );
		} );
	}

	private static IEnumerable<JsonObject> GetLookupTables( MSExcel.Worksheet worksheet )
	{
		var configurationType = worksheet.RangeOrNull<string>( Constants.SpecSheet.RangeNames.SheetType ) switch
		{
			Constants.SpecSheet.SheetTypes.GlobalLookupTables => LookupConfigurationType.GlobalTables,
			Constants.SpecSheet.SheetTypes.GlobalRateTables or Constants.SpecSheet.SheetTypes.ClientRateTables => LookupConfigurationType.RateTables,
			_ => LookupConfigurationType.DataTables // TODO: Is this right?
		};

		var headerOffset = configurationType == LookupConfigurationType.DataTables ? 2 : 1;

		var sheetVersion = worksheet.RangeOrNull<string>( Constants.SpecSheet.RangeNames.SheetVersion );
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

				yield return new JsonObject
				{
					{ "name", tableName },
					{ "version", sheetVersion },
					{ "customizeGlobal", tableInclude.Contains( "/customize" ) },
					{ "columns", new JsonArray().AddItems( data.Rows.First().Select( c => c?.ToString() ) ) },
					{ "rows", data.Rows.Skip( 1 ).Select( r => new JsonArray().AddItems( r.Select( c => c?.ToString() ), includeNulls: true ) ).ToJsonArray() }
				};
			}

			firstColumn = lastColumn.End( DirectionType.ToRight, ignoreEmpty: true );
		}
	}

	private static JsonObject? GetGlobalResourceStrings( MSExcel.Range? resourceTable )
	{
		if ( resourceTable == null ) return null;

		var upperLeft = resourceTable.GetReference();
		var upperRight = upperLeft.End( DirectionType.ToRight );
		var lowerLeft = upperLeft.End( DirectionType.Down );

		var headersRange = upperLeft.Extend( upperRight );
		var dataRange = upperRight.Offset( 1, 0 ).Extend( lowerLeft );

		// 1 row by N columns
		var headers = headersRange.GetArray<string>();
		// X rows by N columns
		var data = dataRange.GetValueArray();
		var rows = data.RowCount;
		var columns = data.ColumnCount;

		return new JsonObject
		{
			{ "version", resourceTable.Worksheet.RangeOrNull<string>( Constants.SpecSheet.RangeNames.SheetVersion ) },
			{ "data", Enumerable.Range( 0, rows )
				.Select( row =>
					new JsonObject
					{
						{ "key", (string?)data[ row, 0 ] }
					}.AddProperties(
						Enumerable.Range( 1, columns - 1 ) // all columns except the key/first
							.Select( col => new JsonKeyProperty( headers[ 0, col ]!, data[ row, col ]?.ToString() ) )
					)
				).ToJsonArray()
			}
		};
	}
}