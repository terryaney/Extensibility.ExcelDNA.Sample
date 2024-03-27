using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Extensibility.Excel.AddIn;

public class WorkbookState
{
	public bool IsGlobalTablesFile { get; init; }
	public bool IsSpecSheetFile { get; init; }
	public bool IsCalcEngine { get; init; }
	public bool IsRTCFile { get; init; }

	public bool IsUploadable { get; init; }
	public bool IsLatestVersion { get; init; }

	public string? UploadedVersion { get; init; }
	public string? CheckedOutBy { get; init; }

	public bool HasRBLeMacro { get; init; }
	public bool HasLinks { get; init; }

	public SheetState SheetState { get; init; } = new();

	public static WorkbookState Current( MSExcel.Application application )
	{
		var activeWorkbook = application.ActiveWorkbook;

		if ( activeWorkbook == null )
		{
			return new();
		}

		var bookNames = activeWorkbook.Names.Cast<MSExcel.Name>().ToArray();
		var isGlobalTablesFile = activeWorkbook.Name.StartsWith( Path.GetFileNameWithoutExtension( Constants.FileNames.GlobalTables ), StringComparison.InvariantCultureIgnoreCase );
		var isRTCFile = activeWorkbook.Name.StartsWith( Path.GetFileNameWithoutExtension( Constants.FileNames.RTCData ), StringComparison.InvariantCultureIgnoreCase );

		var planInfo = !isGlobalTablesFile
			? activeWorkbook
				.Worksheets
				.Cast<MSExcel.Worksheet>()
				.Where( w => w.Name == "Plan Info" )
				.FirstOrDefault()
			: null;

		var isSpecSheet =
			!isGlobalTablesFile && planInfo != null &&
			( 
				planInfo.Names.Cast<MSExcel.Name>().Count( n => n.Name.EndsWith( "!General_Information" ) || n.Name.EndsWith( "!Search_Indexes" ) ) == 2 ||
				bookNames.Count( n => n.Name == "General_Information" || n.Name == "Search_Indexes" ) == 2 
			);

		var isCalcEngine =
			!isSpecSheet && !isGlobalTablesFile &&
			 activeWorkbook.Worksheets.Cast<MSExcel.Worksheet>()
				.Any( s => s.Names.Cast<MSExcel.Name>().Any( n => n.Name.EndsWith( "!SheetType" ) ) && Constants.CalcEngines.SheetTypes.Contains( (string)s.Range[ "SheetType" ].Text ) );

		var hasLinks = activeWorkbook.LinkSources( MSExcel.XlLink.xlExcelLinks ) is Array linkSources && linkSources.Length >= 0;

		var ( liveName, testName ) = GetManagementNames( activeWorkbook!.Name ); // Why does compiler think this is null?

		var liveVersionInfo = isCalcEngine 
			? new
			{
				Description = "// TODO: Need to call api",
				CheckedOutBy = "terry.aney@conduent.com",
				Version = "1.000"
			}
			: null;
		var testVersionInfo = isCalcEngine 
			? new
			{
				Description = "// TODO: Need to call api",
				CheckedOutBy = "terry.aney@conduent.com",
				Version = "1.001"
			}
			: null;

		var isCalcEngineUploadable =
			liveVersionInfo != null &&
			( string.Compare( liveName, activeWorkbook.Name, true ) == 0 || string.Compare( testName, activeWorkbook.Name, true ) == 0 );

		var managementIsLatest =
			liveVersionInfo != null &&
			double.TryParse( activeWorkbook.RangeOrNull<string>( "Version" ), out var currentVersion ) &&
			double.TryParse( ( testVersionInfo ?? liveVersionInfo ).Version, out var managementVersion ) &&
			currentVersion == managementVersion;

		var activeSheet = activeWorkbook.ActiveSheet as MSExcel.Worksheet;

		var rbleMacro = bookNames.FirstOrDefault( n => n.Name == "RBLeMacro" );
		var hasRBLeMacro =
			isCalcEngine && rbleMacro != null &&
			!( (string)rbleMacro.RefersTo ).Contains( "#REF" );

		var sheetNames = activeSheet?.Names.Cast<MSExcel.Name>().ToArray();
		var sheetType = activeSheet?.RangeOrNull<string>( "SheetType" );

		var isGlobalTableSheet = Constants.CalcEngines.GlobalTablesSheetTypes.Contains( sheetType );
		var isXmlMappingSheet = ( sheetNames?.Count( n => n.Name.EndsWith( "!MappingLayouts" ) || n.Name.EndsWith( "!PathToProfileElement" ) || n.Name.EndsWith( "!AuthIdElement" ) ) ?? 0 ) == 3;
		var isUserAccessSheet =
			activeSheet != null &&
			( new[] { "Site Access", "Site Access Delete", "Delete Site Access" }.Contains( activeSheet.Name ) || activeSheet.Name.StartsWith( "Code Tables" ) ) &&
			( sheetNames?.Any( n => n.Name.EndsWith( "!SheetVersion" ) ) ?? false );
		var isExcelJSSheet = ( sheetNames?.Count( n => n.Name.EndsWith( "!Constants" ) || n.Name.EndsWith( "!Inputs" ) || n.Name.EndsWith( "!OutputtedValues" ) ) ?? 0 ) == 3;

		return new()
		{
			CheckedOutBy = liveVersionInfo?.CheckedOutBy,
			IsCalcEngine = isCalcEngine,
			IsGlobalTablesFile = isGlobalTablesFile,
			IsSpecSheetFile = isSpecSheet,
			IsRTCFile = isRTCFile,
			IsUploadable = isCalcEngineUploadable,
			IsLatestVersion = managementIsLatest,
			UploadedVersion = testVersionInfo?.Version ?? liveVersionInfo?.Version,
			HasRBLeMacro = hasRBLeMacro && rbleMacro!.RefersToRange.Worksheet.Name != activeSheet?.Name,
			HasLinks = hasLinks,

			SheetState = new()
			{
				CanExport = isGlobalTablesFile || isUserAccessSheet || isExcelJSSheet,
				CanPreview = isCalcEngine && Constants.CalcEngines.PreviewSheetTypes.Contains( sheetType ),

				IsInputSheet = isCalcEngine && sheetType == Constants.CalcEngines.InputSheetType,
				IsResultSheet = isCalcEngine && Constants.CalcEngines.ResultSheetTypes.Contains( sheetType ),

				IsGlobalTableSheet = isGlobalTableSheet,
				IsXmlMappingSheet = isXmlMappingSheet,
				IsUserAccessSheet = isUserAccessSheet,
				IsExcelJsSheet = isExcelJSSheet,
				IsMacroSheet = hasRBLeMacro && rbleMacro!.RefersToRange.Worksheet.Name == activeSheet?.Name,
			}
		};
	}

	static (string LiveName, string TestName) GetManagementNames( string fileName )
	{
		var managementName =
			Path.GetFileNameWithoutExtension( fileName )
				.Split( ' ' )[ 0 ] // When downloaded by browser, it names it MHA_Spec_Client (#).xls
				.Split( '.' )[ 0 ] // In case they downloaded prior version
				.Replace( "_Test", "" ); // In case they downloaded test from mgmt

		// In case it is saved and client (telegram/browser/etc.) replaced ' ' with _
		var pos = managementName.IndexOf( "_Debug_at_" );
		if ( pos > -1 )
		{
			managementName = managementName[ ..pos ];
		}
		pos = managementName.IndexOf( "_Error_at_" );
		if ( pos > -1 )
		{
			managementName = managementName[ ..pos ];
		}

		managementName += Path.GetExtension( fileName );

		return (managementName, $"{Path.GetFileNameWithoutExtension( managementName )}_Test{Path.GetExtension( managementName )}");
	}
}

public class SheetState
{
	public bool IsResultSheet { get; init; }
	public bool IsInputSheet { get; init; }
	
	public bool CanPreview { get; init; }
	public bool CanExport { get; init; }

	public bool IsGlobalTableSheet { get; init; }
	public bool IsXmlMappingSheet { get; init; }
	public bool IsUserAccessSheet { get; init; }
	public bool IsExcelJsSheet { get; init; }
	public bool IsMacroSheet { get; init; }
}