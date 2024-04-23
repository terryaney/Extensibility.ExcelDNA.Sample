using ExcelDna.Integration;
using ExcelDna.Registration;
using Microsoft.Extensions.Configuration;
using System.Text.Json.Nodes;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public class AddIn : IExcelAddIn
{
	internal static string XllName = null!;
	internal static string XllPath = null!;
	internal static string ResourcePath => Directory.CreateDirectory( Path.Combine( XllPath, "Resources" ) ).FullName;

	internal static AddInSettings Settings = new();
	internal static AddIn CurrentAddin { get; private set; } = null!;
	internal static FileWatcherNotification settingsProcessor = null!;

	public void AutoOpen()
	{
		CurrentAddin = this;

		// Store this for access from anywhere in my workflows: https://groups.google.com/g/exceldna/c/1rScvDdeVOk/m/euij1L-VihoJ
		XllName = (string)XlCall.Excel( XlCall.xlGetName );
		XllPath = Path.GetDirectoryName( XllName )!;

		settingsProcessor = new( 
			notificationDelay: 300, 
			path: XllPath, 
			filter: "appsettings*.json", 
			action: e => {
				try
				{
					IConfiguration configuration = new ConfigurationBuilder()
						.AddJsonFile( Path.Combine( XllPath, "appsettings.json" ), optional: true )
						.AddJsonFile( Path.Combine( XllPath, "appsettings.secrets.json" ), optional: true )
						.Build();

					Settings = configuration.GetSection( "addInSettings" ).Get<AddInSettings>() ?? new();
				}
				catch ( Exception ex )
				{
					Ribbon.LogError( "Unable to initialize IConfiguraiton for appsettings.json.  Using default settings.", ex );
					Settings = new();
				}

				// Don't know why I need QueueAsMacro here.  Without it, Excel wouldn't close gracefully.
				ExcelAsyncUtil.QueueAsMacro( () => Ribbon.CurrentRibbon?.InvalidateFeatures() );
			} 
		);

		settingsProcessor.Changed( "appsettings.json" );

		RegisterFunctions();

		ExcelIntegration.RegisterUnhandledExceptionHandler( UnhandledExceptionHandler );
	}

	public void AutoClose() { }

	/// <summary>
	/// Global exception handler for all unhandled exceptions in ExcelDna functions.
	/// </summary>
	/// <remarks>
	/// This only is triggered for errors from [ExcelFunction] functions in Excel cells (not called in VBA calls to VBAHelpers or errors from Ribbon).
	/// </remarks>
	private object UnhandledExceptionHandler( object exception )
	{
		var caller = ExcelApi.GetCaller();

		// Calculation error is happening on Excel calculation thread, so need to QueueAsMacro to get back to main Excel UI thread.
		// https://groups.google.com/d/msg/exceldna/cHD8Tx56Msg/MdPa2PR13hkJ
		// Explains why needs caller other XlCall methods inside
		ExcelAsyncUtil.QueueAsMacro( () => Ribbon.CurrentRibbon.LogFunctionError( caller, exception ) );

		return ExcelError.ExcelErrorValue;
	}

	/// <summary>
	/// Modify the Excel-DNA function registration, by applying various transformations before the functions are registered.
	/// </summary>
	/// <url>https://github.com/Excel-DNA/Registration</url>
	private static void RegisterFunctions()
	{		
		ExcelRegistration
			.GetExcelFunctions()
			.Select( UpdateHelpTopic )
			.RegisterFunctions();
	}

	private static ExcelFunctionRegistration UpdateHelpTopic( ExcelFunctionRegistration funcReg )
	{
		funcReg.FunctionAttribute.HelpTopic = "http://www.bing.com";
		return funcReg;
	}
}