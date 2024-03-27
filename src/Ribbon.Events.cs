using ExcelDna.Integration.CustomUI;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	string[] RibbonStatesToInvalidateOnWorkbookChange =>
		new[] {
			"exportConfigurations", "processGlobalTables",

			"manageCalcEngine", "debugCalcEngines",

			"navigateToTable", "navigateToInputs", "navigateToInputData", "navigateToCalculationInputs", "navigateToInputTables", "navigateToFrameworkInputs",

			"processWorkbook", "processLocalBatch", "convertToRBLe", 

			"exportxDSData", "exportJsonData", "exportAuditxDSHeaders", "exportMappedxDSData",

			"auditCalcEngineTabs"			
		}.Concat( RibbonStatesToInvalidateOnSheetChange ).Concat( RibbonStatesToInvalidateOnCalcEngineManagement ).ToArray();
	
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
		var showSpecSheet = new Lazy<bool>( () => Convert.ToBase64String( SHA256.HashData( Encoding.UTF8.GetBytes( Features.Salt + "SpecSheet:Allow" ) ) ) == AddIn.Settings.Features.SpecSheet );
		var showGlobalTables = new Lazy<bool>( () => Convert.ToBase64String( SHA256.HashData( Encoding.UTF8.GetBytes( Features.Salt + "GlobalTables:Allow" ) ) ) == AddIn.Settings.Features.GlobalTables );
		var hasxDSDataFields = new Lazy<bool>( () => application.ActiveWorkbook?.Names.Cast<MSExcel.Name>().Any( n => n.Name == "xDSDataFields" ) ?? false );

		return control.Id switch
		{
			"btrRBLe" => showRibbon,

			"debugCalcEngines" => !string.IsNullOrEmpty( AddIn.Settings.SaveHistoryName ),

			"SpecSheet" => showGlobalTables.Value || showSpecSheet.Value,
			"processGlobalTables" => showGlobalTables.Value,
			"navigateToInputs" => !hasxDSDataFields.Value,

			"navigateToInputData" or "navigateToCalculationInputs" or "navigateToFrameworkInputs" => hasxDSDataFields.Value,

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
			"exportxDSData" or "exportJsonData" => !WorkbookState.IsSpecSheetFile && !WorkbookState.IsGlobalTablesFile && !WorkbookState.IsRTCFile && !WorkbookState.IsCalcEngine,

			"convertToRBLe" or "navigateToInputs" or "navigateToInputData"or "navigateToCalculationInputs" or "navigateToFrameworkInputs" 
				or "navigateToInputTables" or "processWorkbook" or "auditCalcEngine" or "auditCalcEngineTab" or "processLocalBatch" => WorkbookState.IsCalcEngine,

			"manageCalcEngine" or "downloadLatestCalcEngine" => WorkbookState.IsCalcEngine && !string.IsNullOrEmpty( WorkbookState.UploadedVersion ),

			"debugCalcEngines" => WorkbookState.IsCalcEngine && !string.IsNullOrEmpty( AddIn.Settings.SaveHistoryName ) && !string.IsNullOrEmpty( WorkbookState.UploadedVersion ),

			"checkInCalcEngine" => WorkbookState.IsCalcEngine && !string.IsNullOrEmpty( WorkbookState.CheckedOutBy ) && string.Compare( WorkbookState.CheckedOutBy, AddIn.Settings.CalcEngineManagement.Email, true ) == 0,
			"checkOutCalcEngine" => WorkbookState.IsCalcEngine && string.Compare( WorkbookState.CheckedOutBy, AddIn.Settings.CalcEngineManagement.Email, true ) != 0 && !string.IsNullOrEmpty( WorkbookState.UploadedVersion ),

			"navigateToTable" => WorkbookState.IsCalcEngine || WorkbookState.IsSpecSheetFile,
			"navigateToBTRCellAddressCell" => WorkbookState.SheetState.IsMacroSheet,
			"navigateToRBLeMacro" => WorkbookState.HasRBLeMacro,
			"linkToLoadedAddIns" => WorkbookState.HasLinks,

			_ => true,
		};
	}

	public string? Ribbon_GetScreentip( IRibbonControl control )
	{
		// TODO: Need to implement
		return control.Id switch
		{
			_ => null,
		};
	}

	public string? Ribbon_GetContent( IRibbonControl control )
	{
		switch ( control.Id )
		{
			case "debugCalcEngines":
			{
				var historyAuthor = AddIn.Settings.SaveHistoryName;
				if ( historyAuthor == "thomas.aney" )
				{
					historyAuthor = "tom.aney";
				}

				var userDirectory = new[] { "terry.aney", "tom.aney" }.Contains( historyAuthor )
					? $"btr.{historyAuthor!.Split( '.' )[ 1 ]}.{historyAuthor.Split( '.' )[ 0 ]}"
					: $"conduent.{historyAuthor!.Split( '.' )[ 1 ]}.{historyAuthor.Split( '.' )[ 0 ]}";

				var ceName = "Conduent_Nexgen_Home_SE.xlsm"; // TODO: workbookState.ManagementName;
				var debugFiles = // TODO: LibraryHelpers.GetDebugCalcEngines( userDirectory, ceName );
					Enumerable.Range( 0, 3 ).Select( i => new { Name = $"{Path.GetFileNameWithoutExtension( ceName )} Debug at {DateTime.Now.AddHours( -1 * ( i + 1 ) ):yyyy-MM-dd hh-mm-sstt} for 011391001.xlsm" } );

				XNamespace ns = "http://schemas.microsoft.com/office/2009/07/customui";
				var menu =
					new XElement( ns + "menu",
						debugFiles.Any()
							? debugFiles.Select( ( f, i ) =>
								new XElement( ns + "button",
									new XAttribute( "id", "managementDownloadFile" + i ),
									new XAttribute( "keytip", i ),
									new XAttribute( "imageMso", "CustomizeXSLTMenu" ),
									new XAttribute( "onAction", "DownloadDebugFile" ),
									new XAttribute( "tag", $"{userDirectory}|{f.Name}" ),
									new XAttribute( "label", f.Name )
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

				return menu.ToString();
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