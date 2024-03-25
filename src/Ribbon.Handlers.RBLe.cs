using ExcelDna.Integration.CustomUI;

namespace KAT.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public void RBLe_BlastEmail( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void RBLe_ShowLog( IRibbonControl? _ )
	{
		ExcelDna.Logging.LogDisplay.Show();
		auditShowLogBadgeCount = 0;
		ribbon.InvalidateControl( "auditShowLog" );
	}

	public void RBLe_OpenHelp( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void RBLe_OpenTemplate( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void RBLe_RefreshRibbon( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void RBLe_HelpAbout( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}
}