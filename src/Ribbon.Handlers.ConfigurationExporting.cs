using ExcelDna.Integration.CustomUI;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public void ConfigurationExporting_ExportWorkbook( IRibbonControl control )
	{
		skipHistoryUpdateOnMoveSpecFromDownloads = true;
		throw new NotImplementedException( "// TODO: Process " + control.Id );
	}

	public void ConfigurationExporting_ProcessGlobalTables( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void ConfigurationExporting_ExportSheet( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}
}