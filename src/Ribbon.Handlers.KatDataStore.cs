using System.Diagnostics;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public async Task KatDataStore_CheckInCalcEngine( IRibbonControl _ )
	{
		await EnsureAddInCredentialsAsync();

		application.Cursor = MSExcel.XlMousePointer.xlWait;
		await apiService.CheckinAsync( WorkbookState.ManagementName, AddIn.Settings.KatUserName, await AddIn.Settings.GetClearPasswordAsync() );
		WorkbookState.CheckInCalcEngine();
		ribbon.Invalidate(); // .InvalidateControls( RibbonStatesToInvalidateOnCalcEngineManagement );
		application.Cursor = MSExcel.XlMousePointer.xlDefault;
	}

	public async Task KatDataStore_CheckOutCalcEngine( IRibbonControl _ )
	{
		await EnsureAddInCredentialsAsync();

		application.Cursor = MSExcel.XlMousePointer.xlWait;
		await apiService.CheckoutAsync( WorkbookState.ManagementName, AddIn.Settings.KatUserName, await AddIn.Settings.GetClearPasswordAsync() );
		WorkbookState.CheckOutCalcEngine();
		ribbon.Invalidate(); // .InvalidateControls( RibbonStatesToInvalidateOnCalcEngineManagement );
		application.Cursor = MSExcel.XlMousePointer.xlDefault;
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

	public async Task KatDataStore_DownloadLatestCalcEngine( IRibbonControl _ ) => await DownloadLatestCalcEngine( WorkbookState.ManagementName );

	private async Task DownloadLatestCalcEngine( string calcEngine, string? destination = null )
	{
		var managedCalcEngine = application.Workbooks.Cast<MSExcel.Workbook>().FirstOrDefault( w => string.Compare( w.Name, calcEngine, true ) == 0 );
		var isDirty = !managedCalcEngine?.Saved ?? false;
		var fullName = Path.Combine( destination ?? Path.GetDirectoryName( ( managedCalcEngine ?? application.ActiveWorkbook ).FullName )!, calcEngine );

		if ( isDirty )
		{
			if ( MessageBox.Show( 
				"You currently have changes in this CalcEngine. If you proceed, all changes will be lost.", 
				"Download Latest Version", 
				MessageBoxButtons.YesNo, 
				MessageBoxIcon.Warning, 
				MessageBoxDefaultButton.Button2 
			) != DialogResult.Yes )
			{
				return;
			}
		}

		managedCalcEngine?.Close( false );

		await EnsureAddInCredentialsAsync();

		application.Cursor = MSExcel.XlMousePointer.xlWait;
		if ( await apiService.DownloadLatestAsync( fullName, AddIn.Settings.KatUserName, await AddIn.Settings.GetClearPasswordAsync() ) )
		{
			// Don't know why I need QueueAsMacro here.  Without it, Excel wouldn't close gracefully.
			ExcelAsyncUtil.QueueAsMacro( () =>
			{
				application.Workbooks.Open( fullName );
				application.Cursor = MSExcel.XlMousePointer.xlDefault;
			} );
		}
		else
		{
			application.Cursor = MSExcel.XlMousePointer.xlDefault;
		}
	}

	public async Task KatDataStore_DownloadDebugFile( IRibbonControl _, string versionKey )
	{
		await EnsureAddInCredentialsAsync();

		application.Cursor = MSExcel.XlMousePointer.xlWait;
		var fileName = await apiService.DownloadDebugAsync( WorkbookState.ManagementName, int.Parse( versionKey ), AddIn.Settings.KatUserName, await AddIn.Settings.GetClearPasswordAsync() );
		if ( !string.IsNullOrEmpty( fileName ) )
		{
			// Don't know why I need QueueAsMacro here.  Without it, Excel wouldn't close gracefully.
			ExcelAsyncUtil.QueueAsMacro( () =>
			{
				application.Workbooks.Open( fileName );
				application.Cursor = MSExcel.XlMousePointer.xlDefault;
			} );
		}
		else
		{
			application.Cursor = MSExcel.XlMousePointer.xlDefault;
		}
	}
}