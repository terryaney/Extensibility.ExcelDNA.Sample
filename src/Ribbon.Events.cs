using ExcelDna.Integration.CustomUI;
using KAT.Camelot.Abstractions.Api.Contracts.Excel.V1.Responses;
using System.Xml.Linq;

namespace KAT.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public bool Ribbon_GetVisible( IRibbonControl control )
	{
		return control.Id switch
		{
			"tabKat" => showRibbon,
			"groupConfigurationExporting" => WorkbookState.ShowDeveloperExports || WorkbookState.ShowGlobalTables,
			"configurationExportingGlobalTables" => WorkbookState.ShowGlobalTables,

			"katDataStoreDebugCalcEnginesMenu" => !string.IsNullOrEmpty( WorkbookState.ManagementName ),
			"katDataStoreCheckOut" => WorkbookState.ShowCalcEngineManagement && WorkbookState.IsCalcEngine && !string.IsNullOrEmpty( AddIn.Settings.CalcEngineManagement.Email ) && string.Compare( WorkbookState.CheckedOutBy, AddIn.Settings.CalcEngineManagement.Email, true ) != 0 && !string.IsNullOrEmpty( WorkbookState.UploadedVersion ),
			"katDataStoreCheckIn" => WorkbookState.ShowCalcEngineManagement && WorkbookState.IsCalcEngine && !string.IsNullOrEmpty( AddIn.Settings.CalcEngineManagement.Email ) && string.Compare( WorkbookState.CheckedOutBy, AddIn.Settings.CalcEngineManagement.Email, true ) == 0,

			"navigationInputs" => !WorkbookState.HasxDSDataFields,
			"navigationInputData" or "navigationCalculationInputs" or "navigationFrameworkInputs" => WorkbookState.HasxDSDataFields,

			_ => true,
		};
	}

	public bool Ribbon_GetEnabled( IRibbonControl control )
	{
		return control.Id switch
		{
			"configurationExportingSheet" => WorkbookState.SheetState.CanExport,
			"configurationExportingWorkbook" => WorkbookState.IsSpecSheetFile || WorkbookState.IsGlobalTablesFile || WorkbookState.IsRTCFile,
			"configurationExportingGlobalTables" => !WorkbookState.IsGlobalTablesFile,

			"katDataStoreManage" => WorkbookState.IsCalcEngine && !string.IsNullOrEmpty( WorkbookState.UploadedVersion ),
			"katDataStoreDownloadLatest" => WorkbookState.IsCalcEngine && !WorkbookState.IsLatestVersion,
			"katDataStoreDebugCalcEnginesMenu" => WorkbookState.IsCalcEngine && !string.IsNullOrEmpty( AddIn.Settings.SaveHistoryName ) && !string.IsNullOrEmpty( WorkbookState.UploadedVersion ),

			"dataExportingDocGenXml" or "dataExportingJsonResultData" or "calcEngineUtilitiesImportBRD" => WorkbookState.SheetState.IsResultSheet,
			"dataExportingExtras" or "dataExportingxDS" or "dataExportingJsonData" => application.ActiveWorkbook != null && !WorkbookState.IsSpecSheetFile && !WorkbookState.IsGlobalTablesFile && !WorkbookState.IsRTCFile && !WorkbookState.IsCalcEngine,
			"dataExportingMappedxDSData" => WorkbookState.ShowDeveloperExports && WorkbookState.SheetState.IsXmlMappingSheet,

			"calcEngineUtilitiesLoadData" => WorkbookState.SheetState.IsInputSheet,
			"calcEngineUtilitiesHighcharts" or "calcEngineUtilitiesPreviewResults" => WorkbookState.SheetState.CanPreview,
			"calcEngineUtilitiesConvertToRBLe" or "calcEngineUtilitiesProcessWorkbook" or "calcEngineUtilitiesLocalBatch" => WorkbookState.IsCalcEngine,
			"calcEngineUtilitiesLinkToLoadedAddIns" => WorkbookState.HasLinks,

			"auditShowDependencies" or "auditHideDependencies" or "auditNonReferencedCells" => application.ActiveWorkbook != null,
			"auditCalcEngine" or "auditCalcEngineTab" or "auditInputResultTabs" => WorkbookState.IsCalcEngine,

			"navigationTable" => WorkbookState.IsCalcEngine || WorkbookState.IsSpecSheetFile,
			"navigationToBTRCellAddressDestination" => WorkbookState.SheetState.IsMacroSheet,
			"navigationToRBLeMacro" => WorkbookState.HasRBLeMacro && !WorkbookState.SheetState.IsMacroSheet,
			"navigationInputs" or "navigationInputData"or "navigationCalculationInputs" or "navigationFrameworkInputs" or "navigationInputTables" => WorkbookState.IsCalcEngine,

			_ => true,
		};
	}

	public string? Ribbon_GetScreentip( IRibbonControl control )
	{
		// TODO: Need to implement
		return control.Id switch
		{
			"katDataStoreDownloadLatest" => $"Download latest version ({WorkbookState.UploadedVersion ?? "Latest"})",
			_ => null,
		};
	}

	public string? Ribbon_GetContent( IRibbonControl control )
	{
		switch ( control.Id )
		{
			case "katDataStoreDebugCalcEnginesMenu":
			{
				string? content = null;
				
				// await ExcelAsyncTaskScheduler.Run( async () => {
				// ExcelAsyncUtil.QueueAsMacro( async () => {
					
					// await EnsureAddInCredentialsAsync();
					EnsureAddInCredentialsAsync().GetAwaiter().GetResult();
					var ceName = WorkbookState.ManagementName;

					var debugFiles = string.IsNullOrEmpty( AddIn.Settings.CalcEngineManagement.Password )
						? Enumerable.Empty<DebugFile>()
						// TODO: await LibraryHelpers.GetDebugCalcEnginesAsync( userDirectory, ceName );
						: Enumerable.Range( 0, 3 ).Select( i => new DebugFile { VersionKey = i, AuthId = "111111111", DateUploaded = DateTime.Now.AddHours( -1 * ( i + 1 ) ) } );

					XNamespace ns = "http://schemas.microsoft.com/office/2009/07/customui";
					var menu =
						new XElement( ns + "menu",
							debugFiles.Any()
								? debugFiles.Select( ( f, i ) =>
									new XElement( ns + "button",
										new XAttribute( "id", "managementDownloadFile" + i ),
										new XAttribute( "keytip", i + 1 ),
										new XAttribute( "imageMso", "CustomizeXSLTMenu" ),
										new XAttribute( "onAction", "Ribbon_OnAction" ),
										new XAttribute( "tag", $"DownloadDebugFile|{f.VersionKey}" ),
										new XAttribute( "label", $"f.AuthId at {( f.DateUploaded.Date == DateTime.Today.Date ? f.DateUploaded.ToShortTimeString() : f.DateUploaded.ToString() )}" )
									)
								)
								: new[] {
										new XElement( ns + "button",
											new XAttribute( "id", "managementDownloadFile0" ),
											new XAttribute( "imageMso", "CustomizeXSLTMenu" ),
											new XAttribute( "label", "No files available" ),
											new XAttribute( "enabled", "false" )
										)
								}
							);

					content = menu.ToString();
				// } );
				return content;
			}
			default: return null;
		}
	}

	private int auditShowLogBadgeCount;
	public Bitmap Ribbon_GetImage( IRibbonControl control )
	{
		switch ( control.Id )
		{
			case "katShowDiagnosticLog":
			{
				using var ms = new MemoryStream( auditShowLogImage );

				var img = System.Drawing.Image.FromStream( ms );

				if ( auditShowLogBadgeCount > 0 )
				{
					var flagGraphics = Graphics.FromImage( img );
					flagGraphics.FillEllipse(
						new SolidBrush( Color.FromArgb( 242, 60, 42 ) ),
						new Rectangle( 11, 0, 19, 19 )
					);
					flagGraphics.DrawString(
						auditShowLogBadgeCount.ToString(),
						new Font( FontFamily.GenericSansSerif, 6, FontStyle.Bold ),
						Brushes.White,
						x: auditShowLogBadgeCount < 10 ? 16 : 13,
						y: 3
					);
				}

				return (Bitmap)img;
			}

			default: throw new ArgumentOutOfRangeException( nameof( control ), $"The id {control.Id} does not support custom image generation." );
		}
	}
}