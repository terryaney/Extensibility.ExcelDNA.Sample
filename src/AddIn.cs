using ExcelDna.Integration;
using ExcelDna.Registration;

namespace KAT.Extensibility.Excel.AddIn;

public class AddIn : IExcelAddIn
{
	internal static string XllPath = null!;

	public void AutoOpen()
	{
		XllPath = Path.GetDirectoryName( (string)XlCall.Excel( XlCall.xlGetName ) )!;

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