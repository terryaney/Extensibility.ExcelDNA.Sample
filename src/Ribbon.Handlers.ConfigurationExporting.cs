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
			ExportSpecificationFile();
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

	void ExportSpecificationFile()
	{
		var owner = new NativeWindow();
		owner.AssignHandle( new IntPtr( application.Hwnd ) );

		var config = GetWindowConfiguration( nameof( ExportSpecification ) );

		var specName = WorkbookState.ManagementName;
		var clientName = Path.GetFileNameWithoutExtension( specName ).Split( '-' ).First( p => !new [] { "MHA", "Spec" }.Contains( p ) );

		var saveLocations =
			AddIn.Settings.SpecificationFileLocations
				.Select( l => l.Replace( "{clientName}", clientName ).Replace( "{specName}", specName ) )
				.ToArray();

		var validLocation = saveLocations.FirstOrDefault( File.Exists );
		using var exportData = new ExportSpecification( 
			validLocation ?? $@"C:\BTR\Camelot\WebSites\Admin\{clientName}\_Developer\{specName}", 
			saveSpecification: validLocation != null,
			config 
		);

		var info = exportData.GetInfo( owner );

		if ( info == null ) return;

		SaveWindowConfiguration( nameof( ExportSpecification ), info.WindowConfiguration );

		MessageBox.Show( "// TODO: Export SpecSheetFile" );
	}
}