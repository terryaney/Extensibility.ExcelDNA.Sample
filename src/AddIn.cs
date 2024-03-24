using ExcelDna.Integration;
using ExcelDna.Registration;

namespace KAT.Extensibility.Excel.AddIn;

public class AddIn : IExcelAddIn
{
	internal static string XllPath = null!;

	public void AutoOpen()
	{
		XllPath = Path.GetDirectoryName( (string)XlCall.Excel( XlCall.xlGetName ) )!;
		
		// TODO: Why am I using this via ExcelAddInExplicitRegistration instead of automatic?
		RegisterFunctions();
	}

	public void AutoClose()
	{
	}

	public void RegisterFunctions()
	{
		ExcelRegistration
			.GetExcelFunctions()
			.Select( UpdateHelpTopic )
			.RegisterFunctions();
	}

	public ExcelFunctionRegistration UpdateHelpTopic( ExcelFunctionRegistration funcReg )
	{
		funcReg.FunctionAttribute.HelpTopic = "http://www.bing.com";
		return funcReg;
	}
}