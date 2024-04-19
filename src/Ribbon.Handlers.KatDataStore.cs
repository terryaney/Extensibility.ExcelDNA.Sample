using ExcelDna.Integration.CustomUI;

namespace KAT.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public async Task KatDataStore_DownloadLatestCalcEngine( IRibbonControl control )
	{
		await EnsureAddInCredentialsAsync();
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public async Task KatDataStore_CheckInCalcEngine( IRibbonControl control )
	{
		await EnsureAddInCredentialsAsync();
		WorkbookState.CheckInCalcEngine();
		ribbon.Invalidate(); // .InvalidateControls( RibbonStatesToInvalidateOnCalcEngineManagement );
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public async Task KatDataStore_CheckOutCalcEngine( IRibbonControl control )
	{
		await EnsureAddInCredentialsAsync();
		WorkbookState.CheckOutCalcEngine();
		ribbon.Invalidate(); // .InvalidateControls( RibbonStatesToInvalidateOnCalcEngineManagement );
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void KatDataStore_ManageCalcEngine( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public async Task KatDataStore_DownloadDebugFile( IRibbonControl control, string versionKey )
	{
		await EnsureAddInCredentialsAsync();
		MessageBox.Show( "// TODO: Process " + control.Id + ", " + versionKey );
	}
}