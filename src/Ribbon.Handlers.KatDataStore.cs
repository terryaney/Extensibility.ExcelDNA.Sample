using System.Diagnostics;
using Aspose.Words.Vba;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace KAT.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public async Task KatDataStore_DownloadLatestCalcEngine( IRibbonControl control )
	{
		await EnsureAddInCredentialsAsync();
		// TODO: Make sure to login when download
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public async Task KatDataStore_CheckInCalcEngine( IRibbonControl _ )
	{
		await EnsureAddInCredentialsAsync();
		await apiService.Checkin( WorkbookState.ManagementName, AddIn.Settings.KatUserName, await AddIn.Settings.GetClearPasswordAsync() );
		WorkbookState.CheckInCalcEngine();
		ribbon.Invalidate(); // .InvalidateControls( RibbonStatesToInvalidateOnCalcEngineManagement );
	}

	public async Task KatDataStore_CheckOutCalcEngine( IRibbonControl _ )
	{
		await EnsureAddInCredentialsAsync();
		await apiService.Checkout( WorkbookState.ManagementName, AddIn.Settings.KatUserName, await AddIn.Settings.GetClearPasswordAsync() );
		WorkbookState.CheckOutCalcEngine();
		ribbon.Invalidate(); // .InvalidateControls( RibbonStatesToInvalidateOnCalcEngineManagement );
	}

	public void KatDataStore_ManageCalcEngine( IRibbonControl _ )
	{
		var url = $"https://btr.lifeatworkportal.com/admin/management/default.aspx?startPage=lnkCalcEngineManagement&startCE={Path.GetFileNameWithoutExtension( WorkbookState.ManagementName )}";

		var psi = new ProcessStartInfo
		{
			FileName = "cmd",
			WindowStyle = ProcessWindowStyle.Hidden,
			UseShellExecute = false,
			RedirectStandardOutput = true,
			// First \"\" is treated as the window title
			Arguments = $"/c start \"\" \"{url}\""
		};
		Process.Start( psi );
	}

	public async Task KatDataStore_DownloadDebugFile( IRibbonControl _, string versionKey )
	{
		await EnsureAddInCredentialsAsync();
		var fileName = await apiService.DownloadDebugAsync( int.Parse( versionKey ), AddIn.Settings.KatUserName, await AddIn.Settings.GetClearPasswordAsync() );
		if ( !string.IsNullOrEmpty( fileName ) )
		{
			// TODO: THrow an exception here and see log displays nicely with one error and debug dropdown still works
			ExcelAsyncUtil.QueueAsMacro( () => application.Workbooks.Open( fileName ) );			
		}
	}
}