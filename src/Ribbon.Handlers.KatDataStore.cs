using ExcelDna.Integration.CustomUI;

namespace KAT.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public void KatDataStore_GetCurrentCalcEngine( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void KatDataStore_CheckInCalcEngine( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void KatDataStore_CheckOutCalcEngine( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void KatDataStore_ManageCalcEngine( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void KatDataStore_DownloadDebugFile( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}
}