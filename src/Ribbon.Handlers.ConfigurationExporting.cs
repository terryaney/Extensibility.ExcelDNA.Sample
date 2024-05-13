using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public void ConfigurationExporting_ExportWorkbook( IRibbonControl _ )
	{
		skipProcessSaveHistory = true;

		if ( WorkbookState.IsGlobalTablesFile )
		{
			ExportGlobalTables();
		}
		else if ( WorkbookState.IsRTCFile )
		{
			ExcelAsyncUtil.QueueAsMacro( () =>
			{
				ExportRtcData();
				MessageBox.Show( "RTC Data Exported" );
			} );
		}
		else if ( WorkbookState.IsSpecSheetFile )
		{
			MessageBox.Show( "// TODO: Export SpecSheetFile" );
		}
	}

	public void ConfigurationExporting_ProcessGlobalTables( IRibbonControl _ )
	{
		var existing = application.GetWorkbook( Constants.FileNames.GlobalTables );
		var downloadName = existing == null
			? Path.Combine( AddIn.ResourcesPath, Constants.FileNames.GlobalTables )
			: null; // Don't download...

		ExportGlobalTables( downloadName );
	}

	public void ConfigurationExporting_ExportSheet( IRibbonControl _ ) =>
		ExportGlobalTables( currentSheet: !WorkbookState.IsGlobalTablesFile );
}