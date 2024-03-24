using ExcelDna.Integration.CustomUI;

namespace KAT.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public void SpecSheet_ExportConfigurations( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void SpecSheet_ProcessGlobalTables( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}
}