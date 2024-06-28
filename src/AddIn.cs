using System.Linq.Expressions;
using System.Reflection;
using ExcelDna.Integration;
using ExcelDna.Registration;
using KAT.Camelot.Extensibility.Excel.AddIn.ExcelApi;
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

		ExcelDna.IntelliSense.IntelliSenseServer.Install();
	}

	public void AutoClose()
	{
		ExcelDna.IntelliSense.IntelliSenseServer.Uninstall();
	}

	/// <summary>
	/// Global exception handler for all unhandled exceptions in ExcelDna functions.
	/// </summary>
	/// <remarks>
	/// This only is triggered for errors from [ExcelFunction] functions in Excel cells or code blocks running fro ExcelAsyncUtil.QueueAsMacro.
	/// </remarks>
	private object UnhandledExceptionHandler( object exception )
	{
		var caller = DnaApplication.GetCaller();

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
		ExcelRegistration
			.GetExcelFunctions()
			.Select( UpdateHelpTopic )
			.ProcessParamsRegistrations()
			.RegisterFunctions();

		var exportedAssemblies = ExcelIntegration.GetExportedAssemblies();

		var debugFunctions = exportedAssemblies
			.SelectMany( a => a.GetTypes() )
			.SelectMany( t => t.GetMethods() )
			.Where( m => m.GetCustomAttribute<KatExcelFunctionAttribute>()?.CreateDebugFunction ?? false )
			.ToList();

		var functionDelegates = 
			debugFunctions
				.Select( m => 
				{
					var parameters = m.GetParameters().Select( p => Expression.Parameter( p.ParameterType ) ).ToArray();

					var call = Expression.Call( null /* static methods */, m, parameters );

					var exceptionParameter = Expression.Parameter( typeof( Exception ), "ex" );

					var body = Expression.TryCatch(
						Expression.Convert( call, typeof( object ) ),
						Expression.Catch( 
							exceptionParameter, 
							Expression.Convert(
								Expression.Property( exceptionParameter, "Message" ),
								typeof( object )
							)
						)
					);
					var lambda = Expression.Lambda( body, parameters ).Compile();

					return lambda;
				} )
				.ToList();

		var functionAttributes = 
			debugFunctions
				.Select( m =>
				{
					var excelFunction = m.GetCustomAttribute<KatExcelFunctionAttribute>()!;
					return new ExcelFunctionAttribute
					{
						Name = $"{m.Name}Debug",
						Description = $"Debug version of {m.Name} that returns value or exception string (instead of #VALUE), see description for {m.Name} for more information.",
						Category = excelFunction.Category,
						IsMacroType = excelFunction.IsMacroType
					};
				} )
				.Cast<object>()
				.ToList();

		var paramAttributes = 
			debugFunctions
				.Select( m => {
					return m.GetParameters()
						.Select( p =>
						{
							var excelArg = p.GetCustomAttribute<ExcelArgumentAttribute>()!;
							return new ExcelArgumentAttribute
							{
								Name = excelArg.Name ?? p.Name,
								Description = excelArg.Description,
								AllowReference = excelArg.AllowReference
							};
						} )
						.Cast<object>()
						.ToList();

				} )
				.Cast<List<object>>()
				.ToList();

		ExcelIntegration.RegisterDelegates(
			functionDelegates,
			functionAttributes,
			paramAttributes
		);
	}

	private static ExcelFunctionRegistration UpdateHelpTopic( ExcelFunctionRegistration funcReg )
	{
		funcReg.FunctionAttribute.HelpTopic = $"https://github.com/terryaney/Documentation.Camelot/blob/main/RBLe/RBLe{funcReg.FunctionAttribute.Category.Replace( " ", "" )}.{funcReg.FunctionAttribute.Name}.md";
		return funcReg;
	}
}