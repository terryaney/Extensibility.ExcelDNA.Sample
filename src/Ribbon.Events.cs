using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using KAT.Camelot.Abstractions.Api.Contracts.Excel.V1.Responses;
using System.Xml.Linq;

namespace KAT.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	string[] RibbonStatesToInvalidateOnFeatureChange => new string[] {
		"btrRBLe", "SpecSheet", "processGlobalTables", "checkInCalcEngine", "checkOutCalcEngine"
	}; 

	string[] RibbonStatesToInvalidateOnWorkbookChange =>
		new[] {
			"exportConfigurations", "processGlobalTables",

			"manageCalcEngine", "debugCalcEngines",

			"navigateToTable", "navigateToInputs", "navigateToInputData", "navigateToCalculationInputs", "navigateToInputTables", "navigateToFrameworkInputs",

			"processWorkbook", "processLocalBatch", "convertToRBLe", 

			"dataExportingExtras", "exportxDSData", "exportAuditxDSHeaders", "exportMappedxDSData",

			"auditExcelCellDepShow", "auditExcelCellDepHide", "auditCalcEngineTabs", "auditCellWithEmptyReferences"
		}
		.Concat( RibbonStatesToInvalidateOnFeatureChange )
		.Concat( RibbonStatesToInvalidateOnSheetChange )
		.Concat( RibbonStatesToInvalidateOnCalcEngineManagement )
		.ToArray();
	
	readonly string[] RibbonStatesToInvalidateOnCalcEngineManagement =
		new[] {
			"downloadLatestCalcEngine", "checkInCalcEngine", "checkOutCalcEngine"
		};
	readonly string[] RibbonStatesToInvalidateOnSheetChange =
		new[] {
			"exportSheet",

			"navigateToBTRCellAddressCell", "navigateToRBLeMacro",

			"loadDataIntoInput", "previewResults", "configureHighCharts", "importBrdSettings",

			"exportRBLDocGen", "exportResultJsonData"
		};

	public bool Ribbon_GetVisible( IRibbonControl control )
	{
		return control.Id switch
		{
			"btrRBLe" => showRibbon,

			"debugCalcEngines" => !string.IsNullOrEmpty( WorkbookState.ManagementName ),

			"SpecSheet" => WorkbookState.ShowDeveloperExports || WorkbookState.ShowGlobalTables,
			"processGlobalTables" => WorkbookState.ShowGlobalTables,
			"navigateToInputs" => !WorkbookState.HasxDSDataFields,

			"navigateToInputData" or "navigateToCalculationInputs" or "navigateToFrameworkInputs" => WorkbookState.HasxDSDataFields,

			"checkOutCalcEngine" => WorkbookState.ShowCalcEngineManagement && ( string.IsNullOrEmpty( WorkbookState.ManagementName ) || ( WorkbookState.IsCalcEngine && string.Compare( WorkbookState.CheckedOutBy, AddIn.Settings.CalcEngineManagement.Email, true ) != 0 && !string.IsNullOrEmpty( WorkbookState.UploadedVersion ) ) ),
			"checkInCalcEngine" => WorkbookState.ShowCalcEngineManagement && WorkbookState.IsCalcEngine && !string.IsNullOrEmpty( WorkbookState.CheckedOutBy ) && string.Compare( WorkbookState.CheckedOutBy, AddIn.Settings.CalcEngineManagement.Email, true ) == 0,

			_ => true,
		};
	}

	public bool Ribbon_GetEnabled( IRibbonControl control )
	{
		return control.Id switch
		{
			// Change when sheet/book changes
			"exportSheet" => WorkbookState.SheetState.CanExport,
			"loadDataIntoInput" => WorkbookState.SheetState.IsInputSheet,
			"exportRBLDocGen" or "exportResultJsonData" or "importBrdSettings" => WorkbookState.SheetState.IsResultSheet,
			"configureHighCharts" or "previewResults" => WorkbookState.SheetState.CanPreview,
			"exportConfigurations" => WorkbookState.IsSpecSheetFile || WorkbookState.IsGlobalTablesFile || WorkbookState.IsRTCFile,
			"processGlobalTables" => WorkbookState.IsGlobalTablesFile,
			"exportMappedxDSData" => WorkbookState.SheetState.IsXmlMappingSheet,
			"dataExportingExtras" or "exportxDSData" or "exportJsonData" => application.ActiveWorkbook != null && !WorkbookState.IsSpecSheetFile && !WorkbookState.IsGlobalTablesFile && !WorkbookState.IsRTCFile && !WorkbookState.IsCalcEngine,

			"auditCalcEngineTabs" => WorkbookState.IsCalcEngine,
			"auditExcelCellDepShow" or "auditExcelCellDepHide" or "auditCellWithEmptyReferences" => application.ActiveWorkbook != null,

			"convertToRBLe" or "navigateToInputs" or "navigateToInputData"or "navigateToCalculationInputs" or "navigateToFrameworkInputs" 
				or "navigateToInputTables" or "processWorkbook" or "auditCalcEngine" or "auditCalcEngineTab" or "processLocalBatch" => WorkbookState.IsCalcEngine,

			"manageCalcEngine" => WorkbookState.IsCalcEngine && !string.IsNullOrEmpty( WorkbookState.UploadedVersion ),
			"downloadLatestCalcEngine" => WorkbookState.IsCalcEngine && !WorkbookState.IsLatestVersion,

			"debugCalcEngines" => WorkbookState.IsCalcEngine && !string.IsNullOrEmpty( AddIn.Settings.SaveHistoryName ) && !string.IsNullOrEmpty( WorkbookState.UploadedVersion ),

			"navigateToTable" => WorkbookState.IsCalcEngine || WorkbookState.IsSpecSheetFile,
			"navigateToBTRCellAddressCell" => WorkbookState.SheetState.IsMacroSheet,
			"navigateToRBLeMacro" => WorkbookState.HasRBLeMacro && !WorkbookState.SheetState.IsMacroSheet,
			"linkToLoadedAddIns" => WorkbookState.HasLinks,

			_ => true,
		};
	}

	public string? Ribbon_GetScreentip( IRibbonControl control )
	{
		// TODO: Need to implement
		return control.Id switch
		{
			"downloadLatestCalcEngine" => $"Download latest version ({WorkbookState.UploadedVersion ?? "Latest"})",
			_ => null,
		};
	}

	public string? Ribbon_GetContent( IRibbonControl control )
	{
		switch ( control.Id )
		{
			case "debugCalcEngines":
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
			case "auditShowLog":
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