using ExcelDna.Integration.CustomUI;
using KAT.Camelot.Abstractions.Api.Contracts.Excel.V1.Responses;
using System.Xml.Linq;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public bool Ribbon_GetVisible( IRibbonControl control )
	{
		try
		{
			return control.Id switch
			{
				"tabKat" => showRibbon,

				"katDataStoreDebugCalcEnginesMenu" => !string.IsNullOrEmpty( WorkbookState.ManagementName ),
				"katDataStoreCheckOut" => WorkbookState.IsCalcEngine && !string.IsNullOrEmpty( AddIn.Settings.KatUserName ) && string.Compare( WorkbookState.CheckedOutBy, AddIn.Settings.KatUserName, true ) != 0 && !string.IsNullOrEmpty( WorkbookState.UploadedVersion ),
				"katDataStoreCheckIn" => WorkbookState.IsCalcEngine && !string.IsNullOrEmpty( AddIn.Settings.KatUserName ) && string.Compare( WorkbookState.CheckedOutBy, AddIn.Settings.KatUserName, true ) == 0,

				"navigationInputs" => !WorkbookState.HasxDSDataFields,
				"navigationInputData" or "navigationCalculationInputs" or "navigationFrameworkInputs" => WorkbookState.HasxDSDataFields,

				_ => true,
			};
		}
		catch ( Exception ex )
		{
			LogError( $"Ribbon_GetVisible {control.Tag}", ex );
			return false;
		}
	}

	public bool Ribbon_GetEnabled( IRibbonControl control )
	{
		try
		{
			return control.Id switch
			{
				"configurationExportingWorkbook" => WorkbookState.IsSpecSheetFile || WorkbookState.IsGlobalTablesFile || WorkbookState.IsRTCFile,
				"configurationExportingSheet" => WorkbookState.SheetState.CanExport,

				"katDataStoreManage" => WorkbookState.IsCalcEngine && !string.IsNullOrEmpty( WorkbookState.UploadedVersion ),
				"katDataStoreDownloadLatest" => WorkbookState.IsCalcEngine && !WorkbookState.IsLatestVersion,
				"katDataStoreDebugCalcEnginesMenu" => WorkbookState.IsCalcEngine && !string.IsNullOrEmpty( AddIn.Settings.SaveHistoryName ) && !string.IsNullOrEmpty( WorkbookState.UploadedVersion ),

				"dataExportingDocGenXml" or "dataExportingJsonResultData" => WorkbookState.SheetState.IsResultSheet,
				"dataExportingExtras" or "dataExportingxDS" or "dataExportingJsonData" => application.ActiveWorkbook != null && !WorkbookState.IsSpecSheetFile && !WorkbookState.IsGlobalTablesFile && !WorkbookState.IsRTCFile && !WorkbookState.IsCalcEngine,
				"dataExportingMappedxDSData" => WorkbookState.SheetState.IsXmlMappingSheet,

				"calcEngineUtilitiesLoadData" => WorkbookState.SheetState.IsInputSheet,
				"calcEngineUtilitiesPreviewResults" => WorkbookState.SheetState.CanPreview,
				"calcEngineUtilitiesConvertToRBLe" or "calcEngineUtilitiesProcessWorkbook" or "calcEngineUtilitiesLocalBatch" => WorkbookState.IsCalcEngine,
				"calcEngineUtilitiesLinkToLoadedAddIns" => WorkbookState.HasLinks,

				"auditShowDependencies" or "auditHideDependencies" or "auditCellsWithEmptyDependencies" => application.ActiveWorkbook != null,
				"auditCalcEngine" or "auditCalcEngineTab" or "auditInputResultTabs" => WorkbookState.IsCalcEngine,

				"navigationTable" => WorkbookState.IsCalcEngine || WorkbookState.IsSpecSheetFile || WorkbookState.SheetState.IsGlobalTableSheet,
				"navigationToBTRCellAddressDestination" => WorkbookState.SheetState.IsMacroSheet,
				"navigationToRBLeMacro" => WorkbookState.HasRBLeMacro && !WorkbookState.SheetState.IsMacroSheet,
				"navigationInputs" or "navigationInputData"or "navigationCalculationInputs" or "navigationFrameworkInputs" or "navigationInputTables" => WorkbookState.IsCalcEngine,

				"katEmailBlast" => !string.IsNullOrEmpty( WorkbookState.SheetState.Name ),

				_ => true,
			};
		}
		catch ( Exception ex )
		{
			LogError( $"Ribbon_GetEnabled {control.Tag}", ex );
			return false;
		}
	}

	public string? Ribbon_GetScreentip( IRibbonControl control )
	{
		try
		{
			return control.Id switch
			{
				"katDataStoreDownloadLatest" => WorkbookState.IsLatestVersion
					? "Download latest version"
					: $"Download latest version ({WorkbookState.CurrentVersion ?? "Current"} -> {WorkbookState.UploadedVersion ?? "Latest"})",
				
				"katDataStoreCheckOut" => string.IsNullOrEmpty( WorkbookState.CheckedOutBy )
					? "Check Out CalcEngine"
					: $"Check Out CalcEngine (checked out by: {WorkbookState.CheckedOutBy})",
				
				_ => null,
			};
		}
		catch ( Exception ex )
		{
			LogError( $"Ribbon_GetGetScreentip {control.Tag}", ex );
			return null;
		}
	}

	public string? Ribbon_GetContent( IRibbonControl control )
	{
		application.Cursor = MSExcel.XlMousePointer.xlWait;
		try
		{
			switch ( control.Id )
			{
				case "katDataStoreDebugCalcEnginesMenu":
				{
					// Deadlock Error :( https://groups.google.com/g/exceldna/c/_pKphutWbvo/m/uvc38llBAAAJ
					// If ExcelDna.Logging.LogDisplay.WriteLine/Show was ever called, the subsequent call to this code would deadlock
					// when I originally had *Async().GetAwaiter().GetResult(). I changed that to Task.Run( () => *Async() ).GetAwaiter().GetResult()
					// and that fixed the deadlock it seems.
					var task = Task.Run( () => GetDebugCalcEnginesAsync() );
					var debugFiles = task.GetAwaiter().GetResult();

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
										new XAttribute( "tag", $"KatDataStore_DownloadDebugFile|{f.VersionKey}" ),
										new XAttribute( "label", GetDebugFileMessage( f ) )
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
		catch ( Exception ex )
		{
			LogError( $"Ribbon_GetContent {control.Tag}", ex );
			return null;
		}
		finally
		{
			application.Cursor = MSExcel.XlMousePointer.xlDefault;
		}
	}

	private static string GetDebugFileMessage( DebugFile debugFile )
	{
		var timeSpan = DateTime.Now - debugFile.DateUploaded;
		var id = debugFile.AuthId;

		return timeSpan switch 
		{
			{ TotalSeconds: < 90 } => $"{id}, {timeSpan.Seconds} seconds ago",
			{ TotalMinutes: < 60 } => $"{id}, {timeSpan.Minutes} minutes ago",
			{ TotalHours: < 24 } => $"{id}, {timeSpan.Hours} hours ago",
			_ => $"{id} on {debugFile.DateUploaded:g}"
		};
	}

	private async Task<IEnumerable<DebugFile>> GetDebugCalcEnginesAsync()
	{
		await EnsureAddInCredentialsAsync();

		var response = await apiService.GetDebugFilesAsync(
			WorkbookState.ManagementName,
			AddIn.Settings.KatUserName,
			await AddIn.Settings.GetClearPasswordAsync()
		);

		if ( response.Validations != null )
		{
			LogValidations( response.Validations );
			return Enumerable.Empty<DebugFile>();
		}

		return response.Response!;
	}

	private int auditShowLogBadgeCount;
	public Bitmap Ribbon_GetImage( IRibbonControl control )
	{
		try
		{
			switch ( control.Id )
			{
				case "katShowDiagnosticLog":
				{
					using var ms = new MemoryStream( auditShowLogImage );

					var img = Image.FromStream( ms );

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
		catch ( Exception ex )
		{
			LogError( $"Ribbon_GetImage {control.Tag}", ex );
			return null!;
		}
	}
}