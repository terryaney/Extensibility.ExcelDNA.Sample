using System.Diagnostics;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public void KatDataStore_CheckInCalcEngine( IRibbonControl _ )
	{
		application.Cursor = MSExcel.XlMousePointer.xlWait;

		RunRibbonTask( async () =>
		{
			try
			{
				await EnsureAddInCredentialsAsync();
				await apiService.CheckinAsync( WorkbookState.ManagementName, AddIn.Settings.KatUserName, await AddIn.Settings.GetClearPasswordAsync() );
				WorkbookState.CheckInCalcEngine();
			}
			finally
			{
				InvalidateRibbon(); // .InvalidateControls( RibbonStatesToInvalidateOnCalcEngineManagement );
			}
		} );
	}

	public void KatDataStore_CheckOutCalcEngine( IRibbonControl _ )
	{
		application.Cursor = MSExcel.XlMousePointer.xlWait;

		RunRibbonTask( async () =>
		{
			try
			{
				await EnsureAddInCredentialsAsync();
				await apiService.CheckoutAsync( WorkbookState.ManagementName, AddIn.Settings.KatUserName, await AddIn.Settings.GetClearPasswordAsync() );
				WorkbookState.CheckOutCalcEngine();
			}
			finally
			{
				InvalidateRibbon(); // .InvalidateControls( RibbonStatesToInvalidateOnCalcEngineManagement );
			}
		} );
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

	public void KatDataStore_DownloadLatestCalcEngine( IRibbonControl _ )
	{
		var fullName = DownloadLatestCalcEngineCheck( WorkbookState.ManagementName );
		RunRibbonTask( () => DownloadLatestCalcEngineAsync( fullName ) );
	}

	public void KatDataStore_DownloadDebugFile( IRibbonControl _, string versionKey )
	{
		application.Cursor = MSExcel.XlMousePointer.xlWait;

		RunRibbonTask( async () => 
		{
			await EnsureAddInCredentialsAsync();
			var fileName = await apiService.DownloadDebugAsync( WorkbookState.ManagementName, int.Parse( versionKey ), AddIn.Settings.KatUserName, await AddIn.Settings.GetClearPasswordAsync() );
			if ( !string.IsNullOrEmpty( fileName ) )
			{
				ExcelAsyncUtil.QueueAsMacro( () => application.Workbooks.Open( fileName ) );
			}
		} );
	}
}