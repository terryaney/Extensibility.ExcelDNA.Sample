using System.Net.Http.Json;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using KAT.Camelot.Abstractions.Api.Contracts.Excel.V1;
using KAT.Camelot.Abstractions.Api.Contracts.Excel.V1.Requests;
using KAT.Camelot.Abstractions.Api.Contracts.Excel.V1.Responses;
using KAT.Camelot.Domain.Extensions;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Extensibility.Excel.AddIn;

public class WorkbookState
{
	private bool isGlobalTablesFile;
	public bool IsGlobalTablesFile => ShowGlobalTables && isGlobalTablesFile;
	public bool IsSpecSheetFile { get; private set; }
	public bool IsCalcEngine { get; private set; }
	private bool isRTCFile;
	public bool IsRTCFile => ShowDeveloperExports && isRTCFile;

	public bool ShowDeveloperExports { get; private set; }
	public bool ShowCalcEngineManagement { get; private set; }
	public bool ShowGlobalTables { get; private set; }

	public string ManagementName { get; private set; } = null!;
	private bool isUploadable;
	public bool IsUploadable => ShowCalcEngineManagement && isUploadable;
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

	public async Task UpdateWorkbookAsync( MSExcel.Workbook activeWorkbook )
	{
		// TODO: Need to check all callers and if only changing sheets, don't need to call api, probably need overload
		// TODO: No more static 'total' create, make methods to initialize diff bits

		if ( activeWorkbook == null )
		{
			ClearState();
			return;
		}

		bookNames = activeWorkbook.Names.Cast<MSExcel.Name>().ToArray();
		isGlobalTablesFile = activeWorkbook.Name.StartsWith( Path.GetFileNameWithoutExtension( Constants.FileNames.GlobalTables ), StringComparison.InvariantCultureIgnoreCase );
		isRTCFile = activeWorkbook.Name.StartsWith( Path.GetFileNameWithoutExtension( Constants.FileNames.RTCData ), StringComparison.InvariantCultureIgnoreCase );

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

		var hasxDSDataFields = isCalcEngine && bookNames.Any( n => n.Name == "xDSDataFields" );

		var hasLinks = activeWorkbook.LinkSources( MSExcel.XlLink.xlExcelLinks ) is Array linkSources && linkSources.Length >= 0;

		var ( liveName, testName ) = GetManagementNames( activeWorkbook.Name );

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

		var calcEngineInfo = await apiService.GetCalcEngineInfoAsync(
			ManagementName,
			AddIn.Settings.KatUserName,
			await AddIn.Settings.GetClearPasswordAsync()
		);

		isUploadable =
			calcEngineInfo != null &&
			( string.Compare( liveName, activeWorkbook.Name, true ) == 0 || string.Compare( testName, activeWorkbook.Name, true ) == 0 );

		CheckedOutBy = calcEngineInfo?.CheckedOutBy;
		UploadedVersion = calcEngineInfo?.Version.ToString();
		CurrentVersion = activeWorkbook.RangeOrNull<string>( "Version" );

		UpdateFeatures();
		UpdateSheet( ( activeWorkbook.ActiveSheet as MSExcel.Worksheet )! );
	}

	public void UpdateVersion( MSExcel.Workbook activeWorkbook ) => UploadedVersion = activeWorkbook.RangeOrNull<string>( "Version" );
	public void CheckInCalcEngine() => CheckedOutBy = null;
	public void CheckOutCalcEngine() => CheckedOutBy = AddIn.Settings.KatUserName;

	public void UpdateFeatures()
	{
		ShowDeveloperExports = Convert.ToBase64String( SHA256.HashData( Encoding.UTF8.GetBytes( Features.Salt + "ShowDeveloperExports:Allow" ) ) ) == AddIn.Settings.Features.ShowDeveloperExports;
		ShowGlobalTables = Convert.ToBase64String( SHA256.HashData( Encoding.UTF8.GetBytes( Features.Salt + "GlobalTables:Allow" ) ) ) == AddIn.Settings.Features.GlobalTables;
		ShowCalcEngineManagement = Convert.ToBase64String( SHA256.HashData( Encoding.UTF8.GetBytes( Features.Salt + "CalcEngineManagement:Allow" ) ) ) == AddIn.Settings.Features.CalcEngineManagement;
	}

	public void UpdateSheet( MSExcel.Worksheet? activeSheet ) => SheetState = new( this, activeSheet );

	private static (string LiveName, string TestName) GetManagementNames( string fileName )
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

	public void ClearState()
	{
		isGlobalTablesFile =
		IsSpecSheetFile =
		IsCalcEngine =
		isRTCFile =
		HasxDSDataFields =
		HasRBLeMacro =
		HasLinks =
		isUploadable = false;

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

	public bool IsResultSheet { get; init; }
	public bool IsInputSheet { get; init; }
	
	public bool CanPreview { get; init; }
	private readonly bool canExport;
	public bool CanExport => ( workbookState?.ShowDeveloperExports ?? false ) && ( workbookState.IsGlobalTablesFile || canExport );

	public bool IsGlobalTableSheet { get; init; }
	public bool IsXmlMappingSheet { get; init; }
	public bool IsUserAccessSheet { get; init; }
	public bool IsExcelJsSheet { get; init; }
	public bool IsMacroSheet { get; init; }

	public SheetState( WorkbookState workbookState, MSExcel.Worksheet? activeSheet )
	{
		this.workbookState = workbookState;

		if ( activeSheet == null )
		{
			return;
		}

		var sheetNames = activeSheet.Names.Cast<MSExcel.Name>().ToArray();
		var sheetType = activeSheet.RangeOrNull<string>( "SheetType" );

		var isGlobalTableSheet = Constants.CalcEngines.GlobalTablesSheetTypes.Contains( sheetType );
		var isXmlMappingSheet = ( sheetNames?.Count( n => n.Name.EndsWith( "!MappingLayouts" ) || n.Name.EndsWith( "!PathToProfileElement" ) || n.Name.EndsWith( "!AuthIdElement" ) ) ?? 0 ) == 3;
		var isUserAccessSheet =
			activeSheet != null &&
			( new[] { "Site Access", "Site Access Delete", "Delete Site Access" }.Contains( activeSheet.Name ) || activeSheet.Name.StartsWith( "Code Tables" ) ) &&
			( sheetNames?.Any( n => n.Name.EndsWith( "!SheetVersion" ) ) ?? false );
		var isExcelJSSheet = ( sheetNames?.Count( n => n.Name.EndsWith( "!Constants" ) || n.Name.EndsWith( "!Inputs" ) || n.Name.EndsWith( "!OutputtedValues" ) ) ?? 0 ) == 3;

		var rbleMacro = workbookState.bookNames.FirstOrDefault( n => n.Name == "RBLeMacro" );

		canExport = isUserAccessSheet || isExcelJSSheet;
		CanPreview = Constants.CalcEngines.PreviewSheetTypes.Contains( sheetType );
		IsInputSheet = sheetType == Constants.CalcEngines.InputSheetType;
		IsResultSheet = Constants.CalcEngines.ResultSheetTypes.Contains( sheetType );

		IsGlobalTableSheet = isGlobalTableSheet;
		IsUserAccessSheet = isUserAccessSheet;
		IsXmlMappingSheet = isXmlMappingSheet;

		IsExcelJsSheet = isExcelJSSheet;
		IsMacroSheet = workbookState.HasRBLeMacro && rbleMacro!.RefersToRange.Worksheet.Name == activeSheet!.Name;
	}
}