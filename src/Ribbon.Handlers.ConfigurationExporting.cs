using System.Text.Json.Nodes;
using System.Xml.Linq;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using KAT.Camelot.Extensibility.Excel.AddIn.RBLe.Dna;
using KAT.Camelot.RBLe.Core;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

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
			var workbook = application.ActiveWorkbook!;
			new ConfigurationExport.Rtc().Export( workbook.Sheets.Cast<MSExcel.Worksheet>() );
			MessageBox.Show( "RTC Data Exported" );
		}
		else if ( WorkbookState.IsSpecSheetFile )
		{
			ExportSpecificationFile();
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

	public void ConfigurationExporting_ExportSheet( IRibbonControl _ ) =>
		ExportGlobalTables( currentSheet: !WorkbookState.IsGlobalTablesFile );

	private void ExportGlobalTables( string? downloadName = null, bool currentSheet = false )
	{
		RunRibbonTask( async () =>
		{
			var katPassword = await AddIn.Settings.GetClearPasswordAsync();

			ExcelAsyncUtil.QueueAsMacro( () => {
				var owner = new NativeWindow();
				owner.AssignHandle( new IntPtr( application.Hwnd ) );

				using var processGlobalTables = new ProcessGlobalTables( 
					currentSheet,
					AddIn.Settings.DataServices,
					GetWindowConfiguration( nameof( ProcessGlobalTables ) )				
				);

				var info = processGlobalTables.GetInfo( 
					AddIn.Settings.KatUserName, 
					katPassword,
					owner
				);

				if ( info == null ) return;

				application.Cursor = MSExcel.XlMousePointer.xlWait;
				application.ScreenUpdating = false;

				SaveWindowConfiguration( nameof( ProcessGlobalTables ), info.WindowConfiguration );

				RunRibbonTask( async () =>
				{
					await UpdateAddInCredentialsAsync( info.UserName, info.Password );
					await DownloadLatestCalcEngineAsync( downloadName );

					SetStatusBar( "Processing Global Tables..." );

					ExcelAsyncUtil.QueueAsMacro( () =>
					{
						try
						{
							var existing = currentSheet ? application.ActiveWorkbook : application.GetWorkbook( Constants.FileNames.GlobalTables )!;

							var sheets = existing.Worksheets.Cast<MSExcel.Worksheet>().ToDictionary( x => x.Name );
							var resourceTable = sheets.TryGetValue( Constants.SpecSheet.TabNames.Localization, out var l )
								? l.RangeOrNull( Constants.SpecSheet.RangeNames.ResourceTable )
								: null;

							var globalSpecifications = ConfigurationExport.GlobalTables.Export( currentSheet
								? new [] { application.ActiveWorksheet() }
								: new [] { sheets[ Constants.SpecSheet.TabNames.DataLookupTables ], sheets[ Constants.SpecSheet.TabNames.RateTables ] }
							);

							RunRibbonTask( async () => {
								var validations = await apiService.UpdateGlobalTablesAsync( info.ClientName, info.Targets, globalSpecifications, info.UserName, info.Password );

								ExcelAsyncUtil.QueueAsMacro( () =>
								{
									try
									{
										if ( downloadName != null )
										{
											existing.Close( SaveChanges: false );
										}

										if ( validations != null )
										{
											ShowValidations( validations );
											return;
										}

										application.ScreenUpdating = true;
										application.Cursor = MSExcel.XlMousePointer.xlDefault;
										MessageBox.Show( owner, $"Successfully updated Global Tables on following environments.{Environment.NewLine + Environment.NewLine + string.Join( ", ", info.Targets )}", "Update Global Tables" );
									}
									catch
									{
										application.ScreenUpdating = true;
										application.Cursor = MSExcel.XlMousePointer.xlDefault;
										throw;
									}
								} );
							}, "ExportGlobalTables (Processing)" );
						}
						catch
						{
							application.Cursor = MSExcel.XlMousePointer.xlDefault;
							application.ScreenUpdating = true;
							throw;
						}
					} );
				}, "ExportGlobalTables (Initialization)" );
			} );
		} );
	}

	private void ExportSpecificationFile()
	{
		var globalTables = application.GetWorkbook( Constants.FileNames.GlobalTables );
		var globalTablesIsOpen = globalTables != null;
		if ( globalTables == null && !File.Exists( Path.Combine( AddIn.ResourcesPath, Constants.FileNames.GlobalTables ) ) )
		{
			MessageBox.Show( "The Global Tables workbook is missing.  Please download it before processing the configuration export.", "Missing Global Tables", MessageBoxButtons.OK, MessageBoxIcon.Warning );
			return;
		}

		var owner = new NativeWindow();
		owner.AssignHandle( new IntPtr( application.Hwnd ) );

		var config = GetWindowConfiguration( nameof( ExportSpecification ) );

		var specName = WorkbookState.ManagementName;
		var clientName = Path.GetFileNameWithoutExtension( specName ).Split( '-' ).First( p => !new [] { "MHA", "Spec" }.Contains( p ) );

		var saveLocations =
			AddIn.Settings.SpecificationFileLocations
				.Select( l => l.Replace( "{clientName}", clientName ).Replace( "{specName}", specName ) )
				.ToArray();

		var validLocation = saveLocations.FirstOrDefault( File.Exists ) ?? $@"C:\BTR\Camelot\WebSites\Admin\{clientName}\_Developer\{specName}";
		using var exportData = new ExportSpecification( 
			validLocation, 
			saveSpecification: validLocation != null,
			config 
		);

		var info = exportData.GetInfo( owner );

		if ( info == null ) return;

		SaveWindowConfiguration( nameof( ExportSpecification ), info.WindowConfiguration );

		skipWorkbookActivateEvents = true;
		var isSaved = application.ActiveWorkbook.Saved;

		if ( !globalTablesIsOpen && File.Exists( Path.Combine( AddIn.ResourcesPath, Constants.FileNames.GlobalTables ) ) )
		{
			application.Workbooks.Open( Path.Combine( AddIn.ResourcesPath, Constants.FileNames.GlobalTables ) );
		}
		XElement globalTablesXml;
		try
		{			
			var globalLookupsConfiguration = new DnaCalcEngineConfigurationFactory( Constants.FileNames.GlobalTables ).LookupsConfiguration;
			globalTablesXml = GetGlobalTablesXml( globalLookupsConfiguration );
		}
		finally
		{
			if ( !globalTablesIsOpen )
			{
				application.GetWorkbook( Constants.FileNames.GlobalTables )!.Close( false );
			}
			skipWorkbookActivateEvents = false;
			application.ActiveWorkbook.Saved = isSaved;
		}


		new ConfigurationExport.Specification().Export( info, application.ActiveWorkbook!, globalTablesXml );

		var currentLocation = application.ActiveWorkbook.FullName;
		if ( info.SaveSpecification && string.Compare( currentLocation, validLocation, true ) != 0 )
		{
			skipProcessSaveHistory = true;
			application.DisplayAlerts = false;

			try
			{
				application.ActiveWorkbook.SaveAs( validLocation );
				try
				{
					if ( !System.Diagnostics.Debugger.IsAttached )
					{
						File.Delete( currentLocation ); // Clean up from temporary location
					}
				}
				catch { }
			}
			finally
			{
				skipProcessSaveHistory = false;
				application.DisplayAlerts = true;
			}
		}

		MessageBox.Show( "Export complete.", "Specification Export", MessageBoxButtons.OK, MessageBoxIcon.Information );
	}
}