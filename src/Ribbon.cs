using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration.Extensibility;
using KAT.Camelot.Domain.Extensions;
using Microsoft.Extensions.DependencyInjection;
using System.Collections.Concurrent;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Extensibility.Excel.AddIn;

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

	private readonly WorkbookState WorkbookState;
	private readonly ApiService apiService;

	public Ribbon()
	{
		application = ( ExcelDnaUtil.Application as MSExcel.Application )!;

		var assembly = Assembly.GetExecutingAssembly();

		using var streamImage = assembly.GetManifestResourceStream( "KAT.Extensibility.Excel.AddIn.Resources.ShowScriptBlockMark.png" )!;
		using var ms = new MemoryStream();
		streamImage.CopyTo( ms );
		auditShowLogImage = ms.ToArray();

		using var streamXml = assembly.GetManifestResourceStream( "KAT.Extensibility.Excel.AddIn.Resources.Ribbon.xml" )!;
		using var sr = new StreamReader( streamXml );
		customUi = sr.ReadToEnd();

		// Create service collection
		var services = new ServiceCollection();
		services.AddHttpClient();
		var serviceProvider = services.BuildServiceProvider();
		var clientFactory = serviceProvider.GetService<IHttpClientFactory>()!;
		apiService = new ApiService( clientFactory );
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

		WorkbookState.UpdateFeatures();
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
		var actionParts = control.Tag.Split( '|' );

		ExcelAsyncUtil.QueueAsMacro( async () =>
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

				if ( mi!.ReturnType == typeof( Task ) )
				{
					await (Task)mi.Invoke( this, new object[] { control }.Concat( parameters ).ToArray() )!;
				}
				else
				{
					mi.Invoke( this, new object[] { control }.Concat( parameters ).ToArray() );
				}
			}
			catch ( Exception ex )
			{
				// I had originally wrapped my call to LogError in *another* QueueAsMacro, but I don't think that is needed.
				// Leaving it here as documenation.  I've found myself sprinkling QueueAsMacro in many/all of my Async ribbon 
				// button events to ensure that Excel closes cleanly.  Unfortunately, I don't have my head wrapped around the 
				// whole async/await and thread context issues.
				LogError( $"Ribbon_OnAction {control.Tag}", ex );
				application.Cursor = MSExcel.XlMousePointer.xlDefault;
				// ExcelAsyncUtil.QueueAsMacro( () => LogError( $"Ribbon_OnAction {control.Tag}", ex ) );
			}
		} );
	}

	internal static void LogError( string message, Exception ex )
	{
		var exDisplay =
			ex.InnerException ?? // Exception in ribbon handler method
			ex; // Exception in try clause above discovering the method to invoke.

		ExcelDna.Logging.LogDisplay.WriteLine( $"{message} Exception: {exDisplay.Message}{Environment.NewLine}{exDisplay.StackTrace}" );

		exDisplay = exDisplay.InnerException;

		while ( exDisplay != null )
		{
			ExcelDna.Logging.LogDisplay.WriteLine( $"Inner Exception: {exDisplay.Message}{Environment.NewLine}Trace: {exDisplay.StackTrace}" );
			exDisplay = exDisplay.InnerException;
		}
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

	private async Task UpdateAddInCredentialsAsync( string userName, string password )
	{
		application.StatusBar = "Saving KAT credentials...";

		// Disable edit notifications...
		AddIn.settingsProcessor.Disable();

		var appSettingsPath = Path.Combine( AddIn.XllPath, "appsettings.json" );
		var appSecretsPath = Path.Combine( AddIn.XllPath, "appsettings.secrets.json" );
		var encryptedPassword = await AddInSettings.EncryptPasswordAsync( password );

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

		AddIn.Settings.SetCredentials( userName, encryptedPassword );
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
}