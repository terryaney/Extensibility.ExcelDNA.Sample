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

	public void KatDataStore_ManageCalcEngine( IRibbonControl _ ) => 
		OpenUrl( $"https://btr.lifeatworkportal.com/admin/management/default.aspx?startPage=lnkCalcEngineManagement&startCE={Path.GetFileNameWithoutExtension( WorkbookState.ManagementName )}" );

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
			var response = await apiService.DownloadDebugAsync( WorkbookState.ManagementName, int.Parse( versionKey ), AddIn.Settings.KatUserName, await AddIn.Settings.GetClearPasswordAsync() );

			ExcelAsyncUtil.QueueAsMacro( () => {
				if ( response.Validations != null )
				{
					LogValidations( response.Validations );
					return;
				}

				application.Workbooks.Open( response.Response! );
			} );
		} );
	}
}