using ExcelDna.Integration;
using ExcelDna.Registration;
using Microsoft.Extensions.Configuration;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public class AddIn : IExcelAddIn
{
	internal static string XllName = null!;
	internal static string XllPath = null!;
	internal static string ResourcesPath => Directory.CreateDirectory( Path.Combine( XllPath, "Resources" ) ).FullName;
	
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
					Ribbon.ShowException( ex, "Unable to initialize IConfiguraiton for appsettings.json.  Using default settings." );
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
	/// This only is triggered for errors from [ExcelFunction] functions in Excel cells or code blocks running fro ExcelAsyncUtil.QueueAsMacro.
	/// </remarks>
	private object UnhandledExceptionHandler( object exception )
	{
		var caller = ExcelApi.GetCaller();

		if ( caller != null )
		{
			// Calculation error is happening on Excel calculation thread, so need to QueueAsMacro to get back to main Excel UI thread.
			// https://groups.google.com/d/msg/exceldna/cHD8Tx56Msg/MdPa2PR13hkJ
			ExcelAsyncUtil.QueueAsMacro( () => Ribbon.CurrentRibbon.LogFunctionError( caller, exception ) );
		}
		else
		{
			Ribbon.ShowException( ( exception as Exception )! );
		}

		return ExcelError.ExcelErrorValue;
	}

	/// <summary>
	/// Modify the Excel-DNA function registration, by applying various transformations before the functions are registered.
	/// </summary>
	/// <url>https://github.com/Excel-DNA/Registration</url>
	private static void RegisterFunctions()
	{		
		// Decided to use this explicit function registration for a couple of reasons...
		//
		// 1) Default parameters don't work correctly and you need to pass as object, then check for ExcelMissing.  I didn't want to include ExcelDNA in the 'functions' repo for the RBLe service, so needed to add a layer.
		// 2) Documentation for Excel (which will be nice to have) needs ExcelDNA library, again, didn't want to share across to RBLe service.
		// 3) There were some 'public' methods that were registering in Excel, that I didn't want there so I used ExcelDNA hidden, and again, that wouldn't compile in RBLe service.
		// 4) Can add a 'debug' version of functions to return object (which can return a number or exception method) without polluting RBLe service code
		//
		// So bit of a pain when you don't need documentation and you have to add the function then add the Excel signature here, but the same thing kind of happens in RBLe when you have to register functions with SSG.
		//
		// Only include C#/Xml documentation if the signature doesn't exist anywhere else (i.e. BTRNumberFormat uses a public method that can not be used by Excel, but can be used by RBLe, so only have signature here)

		ExcelRegistration
			.GetExcelFunctions()
			.Select( UpdateHelpTopic )
			.RegisterFunctions();
	}

	private static object DebugFunction( Func<object> func )
	{
		try
		{
			return func();
		}
		catch ( Exception ex )
		{
			return ex.Message;
		}
	}

	private static ExcelFunctionRegistration UpdateHelpTopic( ExcelFunctionRegistration funcReg )
	{
		// TODO: Ability to run markdown help files locally.
		funcReg.FunctionAttribute.HelpTopic = "http://www.bing.com";
		return funcReg;
	}
}