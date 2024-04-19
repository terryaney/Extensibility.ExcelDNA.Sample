using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration.Extensibility;
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

	private WorkbookState WorkbookState = new();

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

	public async Task InvalidateSettingsAsync()
	{
		RemoveEventHandlers();
		showRibbon = AddIn.Settings.ShowRibbon;
		AddEventHandlers();

		if ( application.ActiveWorkbook != null )
		{
			WorkbookState = await WorkbookState.GetCurrentAsync( application );
		}
		ribbon.InvalidateControls( RibbonStatesToInvalidateOnFeatureChange );
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
		// Need ExcelAsyncUtil.QueueAsMacro so that I can use XlCall.Excel API...some calls work without this (or Application.Run(macroName))
		// hack, but some (ExcelApi.GetText) don't.
		//
		// I couldn't make my functions static to allow the use of Application.Run() workaround because my addin/application event handlers and
		// supporting functions needed access to share variables (i.e. the WorkBookBeforeSave() has a flag - skipHistoryUpdateOnSave - that needs to
		// be toggled off during some of my Ribbon callbacks) and trying to make callbacks static just made code too combersome to maintain.
		//
		// QueueAsMacro Comment: http://stackoverflow.com/questions/31038649/passing-an-excel-range-from-vba-to-c-sharp-via-excel-dna#comment56086701_31047222
		//
		// Application.Run() to enable XlCall.Excel functionality:
		// https://groups.google.com/forum/#!topic/exceldna/YLf6xWfBdQU
		// https://groups.google.com/d/msg/exceldna/t4BDHk_rnQI/S9N1cqQVRw4J
		// https://groups.google.com/forum/#!topic/exceldna/BR5sNFeHvdA
		// application.Run( ctrl.Id, ctrl.Tag );
		// application.Run( control.Tag );

		var actionParts = control.Tag.Split( ':' );

		ExcelAsyncUtil.QueueAsMacro( () =>
		{
			try
			{
				var mi = typeof( Ribbon ).GetMethod( actionParts[ 0 ] )!;
				mi.Invoke( this, new object[] { control }.Concat( actionParts.Skip( 1 ) ).ToArray() );
			}
			catch ( Exception ex )
			{
				LogError( $"Ribbon_OnAction {control.Tag}", ex );
			}
		} );
	}

	private readonly ConcurrentDictionary<string, string?> cellsInError = new();

	private static void LogError( string message, Exception ex )
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
		// If I call LogDisplay.WriteLine *outside* QueueAsMacro, it shows, but if I call
		// it within QueueAsMacro it doesn't show for some reason.
		// Posted to: https://groups.google.com/forum/#!topic/exceldna/97aS22hYR68
		// No response yet.

		var address = caller.GetAddress();
		var formula = caller.GetFormula();

		var reportError = !cellsInError.TryGetValue( address, out var failedFormula ) || failedFormula != formula;
		cellsInError[ address ] = formula;

		if ( reportError )
		{
			var message = $"Error: {address} {formula ?? "unavailable"}{Environment.NewLine}{exception}";
		
			ExcelDna.Logging.LogDisplay.RecordLine( message );

			auditShowLogBadgeCount++;
			ribbon.InvalidateControl( "auditShowLog" );
		}		
	}
}