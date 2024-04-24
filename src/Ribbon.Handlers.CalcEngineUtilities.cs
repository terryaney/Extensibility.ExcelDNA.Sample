﻿using ExcelDna.Integration.CustomUI;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public void CalcEngineUtilities_PopulateInputTab( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void CalcEngineUtilities_ProcessWorkbook( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void CalcEngineUtilities_PreviewResults( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void CalcEngineUtilities_ConfigureHighCharts( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void CalcEngineUtilities_LocalBatchCalc( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void CalcEngineUtilities_DownloadGlobalTables( IRibbonControl _ )
	{
		var fullName = DownloadLatestCalcEngineCheck( Constants.FileNames.GlobalTables, AddIn.ResourcesPath );
		RunRibbonTask( () => DownloadLatestCalcEngineAsync( fullName ) );
	}

	public void CalcEngineUtilities_DownloadHelpersCalcEngine( IRibbonControl _ )
	{
		var fullName = DownloadLatestCalcEngineCheck( Constants.FileNames.Helpers, AddIn.ResourcesPath );
		RunRibbonTask( () => DownloadLatestCalcEngineAsync( fullName ) );
	}

	public void CalcEngineUtilities_ConvertToRBLe( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void CalcEngineUtilities_ImportBrdSettings( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}

	public void CalcEngineUtilities_LinkToLoadedAddIns( IRibbonControl _ ) => UpdateWorkbookLinks( application.ActiveWorkbook );

	private void UpdateWorkbookLinks( MSExcel.Workbook wb )
	{
		if ( wb == null )
		{
			ExcelDna.Logging.LogDisplay.RecordLine( $"LinkToLoadedAddIns: ActiveWorkbook is null." );
			return;
		}

		if ( Path.GetFileName( wb.Name ) != "RBL.Template.xlsx" || !WorkbookState.HasLinks ) return;

		var linkSources = ( wb.LinkSources( MSExcel.XlLink.xlExcelLinks ) as Array )!;

		var protectedInfo = wb.ProtectStructure
			? new[] { "Entire Workbook" }
			: wb.Worksheets.Cast<MSExcel.Worksheet>().Where( w => w.ProtectContents ).Select( w => string.Format( "Worksheet: {0}", w.Name ) ).ToArray();

		if ( protectedInfo.Length > 0 )
		{
			MessageBox.Show( "Unable to update links due to protection.  The following items are protected:\r\n\r\n" + string.Join( "\r\n", protectedInfo ), "Unable to Update", MessageBoxButtons.OK, MessageBoxIcon.Warning );
			return;
		}

		var saved = wb.Saved;

		foreach ( var addin in application.AddIns.Cast<MSExcel.AddIn>().Where( a => a.Installed ) )
		{
			var fullName = addin.FullName;
			var name = Path.GetFileName( fullName );

			foreach ( var o in linkSources )
			{
				var link = (string)o;
				var linkName = Path.GetFileName( link );

				if ( string.Compare( name, linkName, true ) == 0 )
				{
					try
					{
						application.ActiveWorkbook.ChangeLink( link, fullName );
					}
					catch ( Exception ex )
					{
						ExcelDna.Logging.LogDisplay.RecordLine( $"LinkToLoadedAddIns Exception:\r\n\tAddIn Name:{addin.Name}\r\n\tapplication Is Null:{application == null}\r\n\tapplication.ActiveWorkbook Is Null:{application?.ActiveWorkbook == null}\r\n\tName: {name}\r\n\tLink: {link}\r\n\tFullName: {fullName}\r\n\tMessage: {ex.Message}" );
						throw;
					}
				}
			}
		}

		wb.Saved = saved;
	}
}