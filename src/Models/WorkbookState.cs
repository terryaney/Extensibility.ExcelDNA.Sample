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
	public bool IsGlobalTablesFile { get; private set; }
	public bool IsSpecSheetFile { get; private set; }
	public bool IsCalcEngine { get; private set; }
	public bool IsRTCFile { get; private set; }

	public bool ShowDeveloperExports { get; private set; }
	public bool ShowCalcEngineManagement { get; private set; }
	public bool ShowGlobalTables { get; private set; }

	public string ManagementName { get; private set; } = null!;
	public bool IsUploadable { get; private set; }
	public bool IsLatestVersion { get; private set; }

	public string? UploadedVersion { get; private set; }
	public string? CheckedOutBy { get; private set; }

	public bool HasxDSDataFields { get; private set; }
	public bool HasRBLeMacro { get; private set; }
	public bool HasLinks { get; private set; }

	public SheetState SheetState { get; private set; } = new();

	private MSExcel.Name[] bookNames = Array.Empty<MSExcel.Name>();

	public void UpdateSheet( MSExcel.Worksheet? activeSheet )
	{
		if ( activeSheet == null )
		{
			SheetState = new();
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

		var rbleMacro = bookNames.FirstOrDefault( n => n.Name == "RBLeMacro" );

		SheetState = new()
		{
			CanExport = ShowDeveloperExports && ( IsGlobalTablesFile || isUserAccessSheet || isExcelJSSheet ),
			CanPreview = IsCalcEngine && Constants.CalcEngines.PreviewSheetTypes.Contains( sheetType ),

			IsInputSheet = IsCalcEngine && sheetType == Constants.CalcEngines.InputSheetType,
			IsResultSheet = IsCalcEngine && Constants.CalcEngines.ResultSheetTypes.Contains( sheetType ),

			IsGlobalTableSheet = ShowDeveloperExports && isGlobalTableSheet,
			IsXmlMappingSheet = ShowDeveloperExports && isXmlMappingSheet,
			IsUserAccessSheet = ShowDeveloperExports && isUserAccessSheet,
			IsExcelJsSheet = ShowDeveloperExports && isExcelJSSheet,
			IsMacroSheet = HasRBLeMacro && rbleMacro!.RefersToRange.Worksheet.Name == activeSheet.Name,
		};
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

		var showDeveloperExports = Convert.ToBase64String( SHA256.HashData( Encoding.UTF8.GetBytes( Features.Salt + "ShowDeveloperExports:Allow" ) ) ) == AddIn.Settings.Features.ShowDeveloperExports;
		var showGlobalTables = Convert.ToBase64String( SHA256.HashData( Encoding.UTF8.GetBytes( Features.Salt + "GlobalTables:Allow" ) ) ) == AddIn.Settings.Features.GlobalTables;
		var showCalcEngineManagement = Convert.ToBase64String( SHA256.HashData( Encoding.UTF8.GetBytes( Features.Salt + "CalcEngineManagement:Allow" ) ) ) == AddIn.Settings.Features.CalcEngineManagement;

		bookNames = activeWorkbook.Names.Cast<MSExcel.Name>().ToArray();
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

		var hasxDSDataFields = isCalcEngine && bookNames.Any( n => n.Name == "xDSDataFields" );

		var hasLinks = activeWorkbook.LinkSources( MSExcel.XlLink.xlExcelLinks ) is Array linkSources && linkSources.Length >= 0;

		var ( liveName, testName ) = GetManagementNames( activeWorkbook.Name );

		var rbleMacro = bookNames.FirstOrDefault( n => n.Name == "RBLeMacro" );
		var hasRBLeMacro =
			isCalcEngine && rbleMacro != null &&
			!( (string)rbleMacro.RefersTo ).Contains( "#REF" );

		var calcEngineInfo = await GetCalcEngineInfoAsync();

		var isCalcEngineUploadable =
			calcEngineInfo != null &&
			( string.Compare( liveName, activeWorkbook.Name, true ) == 0 || string.Compare( testName, activeWorkbook.Name, true ) == 0 );

		var managementIsLatest =
			calcEngineInfo != null &&
			double.TryParse( activeWorkbook.RangeOrNull<string>( "Version" ), out var currentVersion ) &&
			currentVersion == calcEngineInfo.Version;

		ManagementName = liveName;
		CheckedOutBy = calcEngineInfo?.CheckedOutBy;

		ShowDeveloperExports = showDeveloperExports;
		ShowCalcEngineManagement = showCalcEngineManagement;
		ShowGlobalTables = showGlobalTables;

		IsCalcEngine = isCalcEngine;
		IsGlobalTablesFile = showGlobalTables && /* need showGlobalTables? */ isGlobalTablesFile;
		IsSpecSheetFile = isSpecSheet;
		IsRTCFile = showDeveloperExports && isRTCFile;
		IsUploadable = showCalcEngineManagement && isCalcEngineUploadable;
		IsLatestVersion = managementIsLatest;

		UploadedVersion = calcEngineInfo?.Version.ToString();

		HasxDSDataFields = hasxDSDataFields;
		HasRBLeMacro = hasRBLeMacro;
		HasLinks = hasLinks;

		UpdateSheet( ( activeWorkbook.ActiveSheet as MSExcel.Worksheet )! );
	}

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

	private async Task<CalcEngineInfo?> GetCalcEngineInfoAsync()
	{
		if ( string.IsNullOrEmpty( AddIn.Settings.CalcEngineManagement.Password ) )
		{
			return null;
		}

		var url = $"{AddIn.Settings.ApiEndpoint}{ ApiEndpoints.CalcEngines.Build.Get( Path.GetFileNameWithoutExtension( ManagementName ) )}";

		using var httpClient = new HttpClient();
		using var request = new HttpRequestMessage( HttpMethod.Post, url ) 
		{ 
			Content = new StringContent( JsonSerializer.Serialize( 
				new CalcEngineRequest { 
					Name = ManagementName, 
					Email = AddIn.Settings.CalcEngineManagement.Email!, 
					Password = AddIn.Settings.CalcEngineManagement.Password! 
				} ), 
				Encoding.UTF8, 
				"application/json" 
			) 	
		};

		// TODO: Global error handling...if ensure success status code throws error, excel crashes...get better global handling...
		try
		{
			using var response = await httpClient.SendConduentAsync( request );

			response.EnsureSuccessStatusCode();

			return await response.Content.ReadFromJsonAsync<CalcEngineInfo>();
		}
		catch ( Exception ex )
		{
			throw new ApplicationException( $"Unable to get CalcEngine info from {url}.", ex );
		}
	}

	public void ClearState()
	{
		IsGlobalTablesFile =
		IsSpecSheetFile =
		IsCalcEngine =
		IsRTCFile =
		ShowDeveloperExports =
		ShowCalcEngineManagement =
		ShowGlobalTables =
		HasxDSDataFields =
		HasRBLeMacro =
		HasLinks =
		IsUploadable =
		IsLatestVersion = false;

		ManagementName =
		UploadedVersion =
		CheckedOutBy = null!;

		SheetState = new();
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