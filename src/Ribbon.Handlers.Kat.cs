using ExcelDna.Integration.CustomUI;

namespace KAT.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public void Kat_BlastEmail( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void Kat_ShowLog( IRibbonControl? _ )
	{
		ExcelDna.Logging.LogDisplay.Show();
		auditShowLogBadgeCount = 0;
		ribbon.InvalidateControl( "katShowDiagnosticLog" );
	}

	public void Kat_OpenHelp( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void Kat_OpenTemplate( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public async Task Kat_RefreshRibbon( IRibbonControl _ )
	{
		await EnsureAddInCredentialsAsync();
		await WorkbookState.UpdateWorkbookAsync( application.ActiveWorkbook );
		ribbon.Invalidate();
	}

	public void Kat_HelpAbout( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}
}