using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration.Extensibility;
using Microsoft.Extensions.DependencyInjection;
using System.Collections.Concurrent;
using System.Reflection;
using System.Runtime.InteropServices;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Extensibility.Excel.AddIn;

/// <summary>
/// TODO: Add a description of the Ribbon class here.
/// 
/// Additionally, this class was converted to a partial class for easier maintenance and readability due the the amount
/// of code needed to support all 'ribbon functionality' that is required for the add-in.
/// 
/// 1. The Ribbon.RibbonEvents.cs partial class contains all the events for ribbon elements (i.e. OnLoad, GetVisible, GetEnabled, etc.)
/// 2. The Ribbon.Handlers.*.cs partial class files contain ribbon handlers for each 'group' specified in the CustomUI ribbon xml.  
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
				// Deadlock Error :( https://groups.google.com/g/exceldna/c/_pKphutWbvo/m/uvc38llBAAAJ
				// If ExcelDna.Logging.LogDisplay.WriteLine/Show was called before I called my Ribbon_GetContent
				// handler which uses async code, it would deadlock.  I originally had *Async().GetAwaiter().GetResult().
				// I changed that to Task.Run( () => *Async() ).GetAwaiter().GetResult() and that fixed the deadlock it seems.
				// I had originally wrapped my call to LogError in *another* QueueAsMacro, but I don't think that is needed.
				// Leave it here.  I've found myself sprinkling QueueAsMacro in many/all of my Async ribbon button events
				// to ensure that Excel closes cleanly.  Unfortunately, I don't have my head wrapped around the whole
				// async/await and thread context issues.
				LogError( $"Ribbon_OnAction {control.Tag}", ex );
				// ExcelAsyncUtil.QueueAsMacro( () => LogError( $"Ribbon_OnAction {control.Tag}", ex ) );
			}
		} );
	}

	private readonly ConcurrentDictionary<string, string?> cellsInError = new();

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
}