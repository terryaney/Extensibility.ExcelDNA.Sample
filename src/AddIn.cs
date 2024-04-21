using ExcelDna.Integration;
using ExcelDna.Registration;
using Microsoft.Extensions.Configuration;
using System.Text.Json.Nodes;

namespace KAT.Extensibility.Excel.AddIn;

public class AddIn : IExcelAddIn
{
	internal static string XllPath = null!;
	internal static string ResourcePath => Directory.CreateDirectory( Path.Combine( XllPath, "Resources" ) ).FullName;
	internal static string PreferencesPath => Path.Combine( XllPath, "appsettings.preferences.json" );
	internal static JsonObject Preferences
	{
		get
		{
			var path = PreferencesPath;

			return File.Exists( path )
				? ( JsonNode.Parse( File.ReadAllText( path ) ) as JsonObject )!
				: new JsonObject();
		}
	}

	internal static AddInSettings Settings = new();

	internal static AddIn CurrentAddin { get; private set; } = null!;

	internal static FileWatcherNotification settingsProcessor = null!;


	public void AutoOpen()
	{
		CurrentAddin = this;

		// Store this for access from anywhere in my workflows: https://groups.google.com/g/exceldna/c/1rScvDdeVOk/m/euij1L-VihoJ
		XllPath = Path.GetDirectoryName( (string)XlCall.Excel( XlCall.xlGetName ) )!;

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
					// TODO: Need to log this somewhere...event viewer via Logging?
					Console.WriteLine( ex.ToString() );
					Settings = new();
				}

				// https://groups.google.com/g/exceldna/c/ILgL-dW47A4/m/9HrOyClJAQAJ
				// Excel process not shutting down properly if I didn't wrap this in QueueAsMacro.
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
		// https://groups.google.com/d/msg/exceldna/cHD8Tx56Msg/MdPa2PR13hkJ
		// Explains why needs caller here and QueueAsMacro for other XlCall methods.
		var caller = ExcelApi.GetCaller();

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