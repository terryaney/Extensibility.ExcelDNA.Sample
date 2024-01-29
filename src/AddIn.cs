using ExcelDna.Integration;
using ExcelDna.Registration;

namespace KAT.Extensibility.Excel.AddIn;

public class AddIn : IExcelAddIn
{
	public void AutoOpen()
	{
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