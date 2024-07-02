using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration.Extensibility;
using KAT.Camelot.Data;
using KAT.Camelot.Data.Repositories;
using KAT.Camelot.Domain.Configuration;
using KAT.Camelot.Domain.Extensions;
using KAT.Camelot.Domain.Services;
using KAT.Camelot.Domain.Telemetry;
using KAT.Camelot.Extensibility.Excel.AddIn.ExcelApi;
using KAT.Camelot.Infrastructure.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Localization;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

/// <summary>
/// KAT Add-In ribbon implementation to support Excel ribbon functionality.
/// 
/// This class was converted to a partial class (with file separation based on feature) for easier maintenance and readability due the the amount
/// of code needed to support all 'ribbon functionality' that is required for the add-in.
/// 
/// 1. The Ribbon.Events.cs partial class contains all the events for ribbon elements (i.e. OnLoad, GetVisible, GetEnabled, etc.)
/// 2. The Ribbon.Events.Excel.cs partial class contains all the events for Excel application events (i.e. WorkbookActivated, WorkbookDeactivated, etc.)
/// 3. The Ribbon.Handlers.*.cs partial class files contain ribbon handlers for each 'group' specified in the CustomUI ribbon xml.  
/// </summary>
[ComVisible( true )]
public partial class Ribbon : ExcelRibbon
{
	// Need reference to IRibbonUI so I can change the enable/disable state of buttons and 
	// dynmically update the ribbon (i.e. debug CalcEngine dropdown).  Events are in Ribbon.Events.cs partial class.
	private IRibbonUI ribbon = null!;

	private readonly byte[] auditShowLogImage;
	private readonly string customUi;
	private bool showRibbon;
	private readonly MSExcel.Application application;

	public static Ribbon CurrentRibbon { get; private set; } = null!;

	private readonly ServiceProvider serviceProvider;
	private readonly WorkbookState WorkbookState;
	private readonly ApiService apiService;
	private readonly JwtInfo updatesJwtInfo;
	private readonly RBLe.Dna.DnaCalculationService dnaCalculationService;
	private readonly IEmailService emailService = new FakeEmailService();
	private readonly ITextService textService = new FakeTextService();
	private readonly IConfiguration secretsConfiguration;
	private readonly ICalculationChartBuilder calculationChartBuilder;

	internal static string ExtractStringResource( string resourceName )
	{
		var assembly = Assembly.GetExecutingAssembly();

		using var stream = assembly.GetManifestResourceStream( $"KAT.Camelot.Extensibility.Excel.AddIn.Resources.{resourceName}" )!;
		using var sr = new StreamReader( stream );
		return sr.ReadToEnd();
	}
	internal static byte[] ExtractBinaryResource( string resourceName )
	{
		var assembly = Assembly.GetExecutingAssembly();

		using var stream = assembly.GetManifestResourceStream( $"KAT.Camelot.Extensibility.Excel.AddIn.Resources.{resourceName}" )!;
		using var ms = new MemoryStream();
		stream.CopyTo( ms );
		return ms.ToArray();
	}

	public Ribbon()
	{
		application = ( ExcelDnaUtil.Application as MSExcel.Application )!;

		auditShowLogImage = ExtractBinaryResource( "ShowScriptBlockMark.png" );
		customUi = ExtractStringResource( "Ribbon.xml" );

		// Create service collection
		var csPath = Environment.GetEnvironmentVariable( "CAMELOT_CONFIGURATION_PATH" ) ?? @"C:\BTR\GlobalConfiguration";
        var csEnvironment = Environment.GetEnvironmentVariable( "CAMELOT_SECRETS_ENVIRONMENT" );

		secretsConfiguration = new ConfigurationBuilder()
            .AddJsonFile( Path.Combine( csPath, "Camelot.Secrets.json" ), optional: true, reloadOnChange: true )
            .AddJsonFile( Path.Combine( csPath, $"Camelot.Secrets.Development.json" ), optional: true, reloadOnChange: true )
			.Build();

		var theKeepSection = secretsConfiguration.GetSection( "TheKeep" );

		var services = new ServiceCollection();

		services.AddHttpClient();
		services
			.Configure<Domain.Localization.LocalizationOptions>(
				options =>
				{
					options.AssemblyNames =
						new [] { 
							typeof( Infrastructure.ICamelotMarker ), // Infrastructure namespace
							typeof( Data.ICamelotMarker ), // Data namespace
							typeof( Domain.ICamelotMarker ), // Domain
						}
						.Select( t => new AssemblyName( t.Assembly.FullName! ).Name! )
						.ToArray();
				}
			)
			.AddLocalization(options => options.ResourcesPath = "Resources")
			.AddTransient( typeof( IStringLocalizer<> ), typeof( Domain.Localization.StringLocalizer<> ) );

		services.Configure<TheKeepSettings>( theKeepSection );
		
		serviceProvider = services.BuildServiceProvider();

		var theKeepSettings = serviceProvider.GetRequiredService<IOptionsSnapshot<TheKeepSettings>>();
		var logger = serviceProvider.GetRequiredService<ILogger<CalculationSourceContext>>();
		
		IDbConnectionForge connectionForge = new DbConnectionForge( theKeepSettings );
		IDateTimeService dateTimeService = new DateTimeService();
		IxDSRepository xDSRepository = new xDSRepository( 
			connectionForge, 
			dateTimeService, 
			serviceProvider.GetRequiredService<IStringLocalizer<xDSRepository>>() 
		);

		httpClientFactory = serviceProvider.GetRequiredService<IHttpClientFactory>();
		apiService = new ApiService( httpClientFactory, xDSRepository );
		updatesJwtInfo = theKeepSettings.Value.Jwt.RBLe;
		dnaCalculationService = new RBLe.Dna.DnaCalculationService(
			httpClientFactory, emailService, textService, updatesJwtInfo, logger
		);

		var chartHelper = new ChartHelper( serviceProvider.GetRequiredService<IStringLocalizer<Infrastructure.ICamelotMarker>>() );
		calculationChartBuilder = new HighchartsCalculationChartBuilder( chartHelper );
		WorkbookState = new WorkbookState( apiService );
	}

	public override string GetCustomUI( string RibbonID ) => customUi;

	public void Ribbon_OnLoad( IRibbonUI ribbon ) => this.ribbon = ribbon;

	public override void OnConnection( object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom )
	{
		base.OnConnection( Application, ConnectMode, AddInInst, ref custom );

		CurrentRibbon = this;

		showRibbon = AddIn.Settings.ShowRibbon;

		AddEventHandlers();
	}

	public override void OnDisconnection( ext_DisconnectMode RemoveMode, ref Array custom )
	{
		base.OnDisconnection( RemoveMode, ref custom );

		RemoveEventHandlers();
	}

	public void InvalidateFeatures()
	{
		RemoveEventHandlers();
		showRibbon = AddIn.Settings.ShowRibbon;
		AddEventHandlers();
		ribbon.Invalidate(); // .InvalidateControls( RibbonStatesToInvalidateOnFeatureChange );
	}

	private void AddEventHandlers()
	{
		if ( showRibbon )
		{
			application.WorkbookOpen += Application_WorkbookOpen;
			application.WorkbookBeforeSave += Application_WorkbookBeforeSave;
			application.WorkbookAfterSave += Application_WorkbookAfterSave;
			application.WorkbookActivate += Application_WorkbookActivate;
			application.WorkbookDeactivate += Application_WorkbookDeactivate;
			application.SheetActivate += Application_SheetActivate;

			// Used to remove event handlers to all charts that helped with old 'Excel' chart export 
			// functionality, but SSG does not support that so only use Highcharts/Apex now.
			// application.SheetDeactivate += Application_SheetDeactivate;

			// Used to update 'validation lists' in Tahiti spec sheets when any cell values changed, but no longer use Tahiti, 
			// so disabling for now, but may bring back if 'improve' evolution spec sheet functionality.
			// application.SheetChange += Application_SheetChange;
		}
	}

	private void RemoveEventHandlers()
	{
		if ( showRibbon )
		{
			application.WorkbookOpen -= Application_WorkbookOpen;
			application.WorkbookBeforeSave -= Application_WorkbookBeforeSave;
			application.WorkbookAfterSave -= Application_WorkbookAfterSave;
			application.WorkbookActivate -= Application_WorkbookActivate;
			application.WorkbookDeactivate -= Application_WorkbookDeactivate;
			application.SheetActivate -= Application_SheetActivate;
			// application.SheetDeactivate -= Application_SheetDeactivate;
			// application.SheetChange -= Application_SheetChange;
		}
	}

	public void Ribbon_OnAction( IRibbonControl control )
	{
		var tag = control.Tag;
		var actionParts = tag.Split( '|' );

		ExcelAsyncUtil.QueueAsMacro( () =>
		{
			try
			{
				var parameters = actionParts.Skip( 1 ).ToArray();
				var parameterTypes = parameters.Any()
					? new[] { typeof( IRibbonControl ) }.Concat( parameters.Select( p => typeof( string ) ) ).ToArray()
					: null;

				var mi = parameters.Any()
					? typeof( Ribbon ).GetMethod( actionParts[ 0 ], parameterTypes! )
					: typeof( Ribbon ).GetMethod( actionParts[ 0 ] );

				// WARNING: If mi is async Task, need to wrap entire implementation inside Try/Catch and display own exception
				mi!.Invoke( this, new object[] { control }.Concat( parameters ).ToArray() );
			}
			catch ( Exception ex )
			{
				ShowException( ex, $"Ribbon_OnAction {tag}" );
			}
			finally
			{
				// TODO: May have to remove this...if mi is async Task, then this will flip to default and I'm not sure of timing (if mi sets to wait)
				application.Cursor = MSExcel.XlMousePointer.xlDefault;
			}
		} );
	}

	private void RunRibbonTask( Func<Task> action, [CallerMemberName] string actionName = "" )
	{
		Task.Run( async () =>
		{
			try
			{
				await action();
			}
			catch ( Exception ex )
			{
				ShowException( ex, actionName );
			}
			finally
			{
				ExcelAsyncUtil.QueueAsMacro( () => application.Cursor = MSExcel.XlMousePointer.xlDefault );
			}
		} );
	}

	internal static void ShowValidations( ApiValidation[] validations )
	{
		LogValidations( validations );
		ExcelDna.Logging.LogDisplay.Show();
	}

	internal static void LogValidations( ApiValidation[] validations )
	{
		foreach ( var validation in validations )
		{
			ExcelDna.Logging.LogDisplay.WriteLine( $"{validation.Name}: {validation.Message}" );
		}
	}

	internal static void ShowException( Exception ex, string? message = null, IEnumerable<string>? additionalMessages = null )
	{
		var exDisplay = ex; 

		ExcelDna.Logging.LogDisplay.WriteLine( $"{message} Exception: {exDisplay.Message}{Environment.NewLine}{exDisplay.StackTrace}" );

		exDisplay = exDisplay.InnerException;

		while ( exDisplay != null )
		{
			ExcelDna.Logging.LogDisplay.WriteLine( $"Inner Exception: {exDisplay.Message}{Environment.NewLine}Trace: {exDisplay.StackTrace}" );
			exDisplay = exDisplay.InnerException;
		}

		if ( additionalMessages != null )
		{
			foreach ( var additionalMessage in additionalMessages )
			{
				ExcelDna.Logging.LogDisplay.WriteLine( additionalMessage );
			}
		}

		ExcelDna.Logging.LogDisplay.Show();
	}

	private readonly ConcurrentDictionary<string, string?> cellsInError = new();
	public void LogFunctionError( ExcelReference caller, object exception )
	{
		var address = caller.GetAddress();
		var formula = caller.GetFormula();

		var reportError = !cellsInError.TryGetValue( address, out var failedFormula ) || failedFormula != formula;
		cellsInError[ address ] = formula;

		if ( reportError )
		{
			ExcelDna.Logging.LogDisplay.RecordLine( $"Error: {address} {formula ?? "unavailable"}{Environment.NewLine}{exception}" );

			auditShowLogBadgeCount++;
			ribbon.InvalidateControl( "katShowDiagnosticLog" );
		}
	}

	private async Task EnsureAddInCredentialsAsync()
	{
		if ( string.IsNullOrEmpty( AddIn.Settings.KatUserName ) || string.IsNullOrEmpty( AddIn.Settings.KatPassword ) )
		{
			using var credentials = new Credentials( GetWindowConfiguration( nameof( Credentials ) ) );
			
			var info = credentials.GetInfo(  
				AddIn.Settings.KatUserName, 
				await AddIn.Settings.GetClearPasswordAsync() 
			);

			if ( info != null )
			{
				await UpdateAddInCredentialsAsync( info.UserName, info.Password );
				SaveWindowConfiguration( nameof( Credentials ), info.WindowConfiguration );
			}
		}
	}

	private async Task UpdateAddInCredentialsAsync( string userName, string password )
	{
		SetStatusBar( "Saving KAT credentials..." );

		try
		{
			if ( userName != AddIn.Settings.KatUserName || password != await AddIn.Settings.GetClearPasswordAsync() )
			{
				// Disable edit notifications...
				AddIn.settingsProcessor.Disable();

				var appSettingsPath = Path.Combine( AddIn.XllPath, "appsettings.json" );
				var appSecretsPath = Path.Combine( AddIn.XllPath, "appsettings.secrets.json" );
				var encryptedPassword = await AddIn.Settings.SetCredentialsAsync( userName, password );

				static void updateSetting( string path, string key, string value )
				{
					var appSettings = File.Exists( path )
						? ( JsonNode.Parse( File.ReadAllText( path ) ) as JsonObject )!
						: new JsonObject();

					var addInSettings = ( ( appSettings[ "addInSettings" ] ?? appSettings.AddOrUpdate( "addInSettings", new JsonObject() ) ) as JsonObject )!;
					addInSettings.AddOrUpdate( key, value );
					appSettings.Save( path );
				}

				updateSetting( appSettingsPath, "katUserName", userName );
				updateSetting( appSecretsPath, "katPassword", encryptedPassword! );

				AddIn.settingsProcessor.Enable();
			}
		}
		finally
		{
			ClearStatusBar();
		}
	}

	private static void SaveWindowConfiguration( string name, JsonObject windowConfiguration )
	{
		var appSettingsPath = Path.Combine( AddIn.XllPath, "appsettings.json" );

		var appSettings = File.Exists( appSettingsPath )
			? ( JsonNode.Parse( File.ReadAllText( appSettingsPath ) ) as JsonObject )!
			: new JsonObject();

		var windowSettings = ( ( appSettings[ "windowSettings" ] ?? appSettings.AddOrUpdate( "windowSettings", new JsonObject() ) ) as JsonObject )!;
		windowSettings[ name ] = windowConfiguration.Clone();

		// Disable edit notifications...
		AddIn.settingsProcessor.Disable();
		appSettings.Save( appSettingsPath );
		AddIn.settingsProcessor.Enable();
	}

	private static JsonObject? GetWindowConfiguration( string name )
	{
		var appSettingsPath = Path.Combine( AddIn.XllPath, "appsettings.json" );

		var appSettings = File.Exists( appSettingsPath )
			? ( JsonNode.Parse( File.ReadAllText( appSettingsPath ) ) as JsonObject )!
			: new JsonObject();

		return appSettings[ "windowSettings" ]?[ name ] as JsonObject;
	}

	private bool isSpreadsheetGearLicensed;
	private readonly IHttpClientFactory httpClientFactory;

	private async Task<bool> EnsureSpreadsheetGearLicenseAsync()
	{
		if ( isSpreadsheetGearLicensed ) return true;

		SetStatusBar( "Checking for SpreadsheetGear License..." );

		var response = await apiService.GetSpreadsheetGearLicenseAsync( AddIn.Settings.KatUserName, await AddIn.Settings.GetClearPasswordAsync() );

		if ( response.Validations != null )
		{
			ExcelAsyncUtil.QueueAsMacro( () =>
			{
				LogValidations( response.Validations );
				MessageBox.Show( "KAT dependent license not found.", "KAT License", MessageBoxButtons.OK, MessageBoxIcon.Error );
			} );
			return false;
		}

		SpreadsheetGear.Factory.SetSignedLicense( response.Response! );
		return isSpreadsheetGearLicensed = true;
	}

	private void SetStatusBar( string message ) => ExcelAsyncUtil.QueueAsMacro( () => application.StatusBar = $"KAT: {message}" );
	private void ClearStatusBar() => ExcelAsyncUtil.QueueAsMacro( () => {
		if ( ( (string?)application.StatusBar ?? "" ).StartsWith( "KAT: " ) )
		{
			application.StatusBar = "";
			application.Cursor = MSExcel.XlMousePointer.xlDefault;
		}
	} );
	private void InvalidateRibbon() => ExcelAsyncUtil.QueueAsMacro( () => {
		ribbon.Invalidate();
		application.Cursor = MSExcel.XlMousePointer.xlDefault;
	} );

	private string? DownloadLatestCalcEngineCheck( string calcEngine, string? destination = null )
	{
		var managedCalcEngine = application.GetWorkbook( calcEngine );
		var isDirty = !managedCalcEngine?.Saved ?? false;
		var fullName = Path.Combine( destination ?? Path.GetDirectoryName( ( managedCalcEngine ?? application.ActiveWorkbook ).FullName )!, calcEngine );

		if ( 
			isDirty && 
			MessageBox.Show( 
				"You currently have changes in this CalcEngine. If you proceed, all changes will be lost.", 
				"Download Latest Version", 
				MessageBoxButtons.YesNo, 
				MessageBoxIcon.Warning, 
				MessageBoxDefaultButton.Button2 
			) != DialogResult.Yes 
		)
		{
			return null;
		}

		application.Cursor = MSExcel.XlMousePointer.xlWait;
		managedCalcEngine?.Close( false );
		return fullName;
	}

	private async Task DownloadLatestCalcEngineAsync( string? fullName, bool openFile = true )
	{
		if ( string.IsNullOrEmpty( fullName ) ) return;
		
		await EnsureAddInCredentialsAsync();

		var response = await apiService.DownloadLatestAsync( fullName, AddIn.Settings.KatUserName, await AddIn.Settings.GetClearPasswordAsync() );

		if ( response != null )
		{
			ShowValidations( response );
			return;
		}

		if ( openFile )
		{
			ExcelAsyncUtil.QueueAsMacro( () => application.Workbooks.Open( fullName ) );
		}
	}

	private static void OpenUrl( string url )
	{
		try
		{
			var psi = new ProcessStartInfo
			{
				FileName = "cmd",
				WindowStyle = ProcessWindowStyle.Hidden,
				UseShellExecute = false,
				RedirectStandardOutput = true,
				// First \"\" is treated as the window title
				Arguments = $"/c start \"\" \"{url}\""
			};
			Process.Start( psi );
		}
		catch ( Exception ex )
		{
			ShowException( ex, $"Couldn't open {url}." );
		}
	}
}