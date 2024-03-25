using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	private void Application_WorkbookOpen( MSExcel.Workbook Wb )
	{
		Console.WriteLine( "// TODO: Application_WorkbookOpen" );

		// Clear error info whenever a new workbook is opened.  Currenly, only show any 
		// errors after a cell is calculated.  Could call application.Calculate() to force everything
		// to re-evaluate, but that could be expensive, so for now, not doing it, the function log display
		// is just helpful information for CalcEngine developer to 'clean' up their formulas.
		auditShowLogBadgeCount = 0;
		cellsInError.Clear();
		ExcelDna.Logging.LogDisplay.Clear();
	}

	private void Application_WorkbookBeforeSave( MSExcel.Workbook Wb, bool SaveAsUI, ref bool Cancel )
	{
		Console.WriteLine( "// TODO: Application_WorkbookBeforeSave" );

		if ( auditShowLogBadgeCount > 0 )
		{
			RBLe_ShowLog( null );
			Cancel = true;
		}
	}

	private async void Application_WorkbookAfterSave( MSExcel.Workbook Wb, bool Success )
	{
		Console.WriteLine( "// TODO: Application_WorkbookAfterSave" );
		await Task.Delay( 1 );
	}

	private void Application_WorkbookActivate( MSExcel.Workbook Wb )
	{
		Console.WriteLine( "// TODO: Application_WorkbookActivate" );
	}

	private void Application_WorkbookDeactivate( MSExcel.Workbook Wb )
	{
		Console.WriteLine( "// TODO: Application_WorkbookDeactivate" );
	}

	private void Application_SheetActivate( object sheet )
	{
		Console.WriteLine( "// TODO: Application_SheetActivate" );
	}

	private void Application_SheetDeactivate( object sheet )
	{
		Console.WriteLine( "// TODO: Application_SheetDeactivate" );
	}

	private void Application_SheetChange( object sheet, MSExcel.Range target )
	{
		Console.WriteLine( "// TODO: Application_SheetChange" );		
	}
}