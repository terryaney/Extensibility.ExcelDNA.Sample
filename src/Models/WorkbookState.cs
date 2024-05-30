using KAT.Camelot.Abstractions.Api.Contracts.Excel.V1.Responses;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public class WorkbookState
{
	public bool IsGlobalTablesFile { get; private set; }
	public bool IsSpecSheetFile { get; private set; }
	public bool IsCalcEngine { get; private set; }
	public bool IsRTCFile { get; private set; }

	public string ManagementName { get; private set; } = null!;
	public string TestManagementName => $"{Path.GetFileNameWithoutExtension( ManagementName )}_Test{Path.GetExtension( ManagementName )}";

	public bool IsUploadable { get; private set; }
	public bool IsLatestVersion => 
		!string.IsNullOrEmpty( UploadedVersion ) && !string.IsNullOrEmpty( CurrentVersion ) &&
		double.TryParse( CurrentVersion, out var version ) &&
		version == double.Parse( UploadedVersion );

	public string? CurrentVersion { get; private set; }
	public string? UploadedVersion { get; private set; }
	public string? CheckedOutBy { get; private set; }

	public bool HasxDSDataFields { get; private set; }
	public bool HasRBLeMacro { get; private set; }
	public bool HasLinks { get; private set; }

	public SheetState SheetState { get; private set; }

	internal MSExcel.Name[] bookNames = Array.Empty<MSExcel.Name>();
	private readonly ApiService apiService;

	public WorkbookState( ApiService dataLockerService )
	{
		SheetState = new( this, null );
		this.apiService = dataLockerService;
	}

	public void UpdateWorkbook( MSExcel.Workbook? activeWorkbook )
	{
		if ( activeWorkbook == null )
		{
			ClearState();
			return;
		}

		bookNames = activeWorkbook.Names.Cast<MSExcel.Name>().ToArray();
		IsGlobalTablesFile = activeWorkbook.Name.StartsWith( Path.GetFileNameWithoutExtension( Constants.FileNames.GlobalTables ), StringComparison.InvariantCultureIgnoreCase );
		IsRTCFile = activeWorkbook.Name.StartsWith( Path.GetFileNameWithoutExtension( Constants.FileNames.RTCData ), StringComparison.InvariantCultureIgnoreCase );

		var planInfo = !IsGlobalTablesFile
			? activeWorkbook
				.Worksheets
				.Cast<MSExcel.Worksheet>()
				.Where( w => w.Name == "Plan Info" )
				.FirstOrDefault()
			: null;

		var isSpecSheet =
			!IsGlobalTablesFile && planInfo != null &&
			( 
				planInfo.Names.Cast<MSExcel.Name>().Count( n => n.Name.EndsWith( "!General_Information" ) || n.Name.EndsWith( "!Search_Indexes" ) ) == 2 ||
				bookNames.Count( n => n.Name == "General_Information" || n.Name == "Search_Indexes" ) == 2 
			);

		var isCalcEngine =
			!isSpecSheet && !IsGlobalTablesFile &&
			 activeWorkbook.Worksheets.Cast<MSExcel.Worksheet>()
				.Any( s => s.Names.Cast<MSExcel.Name>().Any( n => n.Name.EndsWith( "!SheetType" ) ) && Constants.CalcEngines.IsRBLeSheet( (string)s.Range[ "SheetType" ].Text ) );

		var hasxDSDataFields = isCalcEngine && bookNames.Any( n => n.Name == "xDSDataFields" );

		var hasLinks = activeWorkbook.LinkSources( MSExcel.XlLink.xlExcelLinks ) is Array linkSources && linkSources.Length >= 0;

		var liveName = GetManagementName( activeWorkbook.Name );

		var rbleMacro = bookNames.FirstOrDefault( n => n.Name == "RBLeMacro" );
		var hasRBLeMacro =
			isCalcEngine && rbleMacro != null &&
			!( (string)rbleMacro.RefersTo ).Contains( "#REF" );

		ManagementName = liveName;
		IsCalcEngine = isCalcEngine;
		IsSpecSheetFile = isSpecSheet;
		HasxDSDataFields = hasxDSDataFields;
		HasRBLeMacro = hasRBLeMacro;
		HasLinks = hasLinks;
		CurrentVersion = activeWorkbook.RangeOrNull<string>( "Version" );

		UpdateSheet( ( activeWorkbook.ActiveSheet as MSExcel.Worksheet )! );
	}

	public async Task<ApiValidation[]?> UpdateCalcEngineInfoAsync( string activeWorkbookName )
	{
		if ( !IsCalcEngine )
		{
			IsUploadable = false;
			return null;
		}

		var response = await apiService.GetCalcEngineInfoAsync(
			ManagementName,
			AddIn.Settings.KatUserName,
			await AddIn.Settings.GetClearPasswordAsync()
		);

		if ( response.Validations != null )
		{
			return response.Validations;
		}

		// Response is null when CE Not found in mgmt site...
		var calcEngineInfo = response.Response ?? new CalcEngineInfo { CheckedOutBy = null, Version = 0 };

		IsUploadable = string.Compare( ManagementName, activeWorkbookName, true ) == 0 || string.Compare( TestManagementName, activeWorkbookName, true ) == 0;
		CheckedOutBy = calcEngineInfo.CheckedOutBy;
		UploadedVersion = calcEngineInfo.Version.ToString();

		return null;
	}

	public void UpdateVersion( MSExcel.Workbook activeWorkbook ) => UploadedVersion = activeWorkbook.RangeOrNull<string>( "Version" );
	public void CheckInCalcEngine() => CheckedOutBy = null;
	public void CheckOutCalcEngine() => CheckedOutBy = AddIn.Settings.KatUserName;

	public void UpdateSheet( MSExcel.Worksheet? activeSheet ) => SheetState = new( this, activeSheet );

	private static string GetManagementName( string fileName )
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

		return managementName;
	}

	public void ClearState()
	{
		IsGlobalTablesFile =
		IsSpecSheetFile =
		IsCalcEngine =
		IsRTCFile =
		HasxDSDataFields =
		HasRBLeMacro =
		HasLinks =
		IsUploadable = false;

		ManagementName =
		UploadedVersion =
		CurrentVersion =
		CheckedOutBy = null!;

		SheetState = new( this, null );
	}
}

public class SheetState
{
	private readonly WorkbookState workbookState;

	public string? Name { get; init; }

	public string? SheetType { get; init; }
	public bool IsResultSheet { get; init; }
	public bool IsInputSheet { get; init; }
	
	public bool CanPreview { get; init; }
	private readonly bool canExport;
	public bool CanExport => workbookState.IsGlobalTablesFile || canExport;

	public bool IsGlobalTableSheet { get; init; }
	public bool IsXmlMappingSheet { get; init; }
	public bool IsExcelJsSheet { get; init; }
	public bool IsMacroSheet { get; init; }

	public SheetState( WorkbookState workbookState, MSExcel.Worksheet? activeSheet )
	{
		this.workbookState = workbookState;

		if ( activeSheet == null )
		{
			return;
		}

		Name = activeSheet.Name;
		
		var sheetNames = activeSheet.Names.Cast<MSExcel.Name>().ToArray();
		SheetType = activeSheet.RangeOrNull<string>( "SheetType" );

		var isGlobalTableSheet = Constants.SpecSheet.IsGlobalTablesSheet( SheetType );
		var isXmlMappingSheet = ( sheetNames?.Count( n => n.Name.EndsWith( "!MappingLayouts" ) || n.Name.EndsWith( "!PathToProfileElement" ) || n.Name.EndsWith( "!AuthIdElement" ) ) ?? 0 ) == 3;
		var isExcelJSSheet = ( sheetNames?.Count( n => n.Name.EndsWith( "!Constants" ) || n.Name.EndsWith( "!Inputs" ) || n.Name.EndsWith( "!OutputtedValues" ) ) ?? 0 ) == 3;

		var rbleMacro = workbookState.bookNames.FirstOrDefault( n => n.Name == "RBLeMacro" );

		canExport = isExcelJSSheet;
		CanPreview = Constants.CalcEngines.IsPreviewSheet( SheetType );
		IsInputSheet = SheetType == Constants.CalcEngines.SheetTypes.Input;
		IsResultSheet = Constants.CalcEngines.IsResultSheet( SheetType );

		IsGlobalTableSheet = isGlobalTableSheet;
		IsXmlMappingSheet = isXmlMappingSheet;

		IsExcelJsSheet = isExcelJSSheet;
		IsMacroSheet = workbookState.HasRBLeMacro && rbleMacro!.RefersToRange.Worksheet.Name == activeSheet!.Name;
	}
}