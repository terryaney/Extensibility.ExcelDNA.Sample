using ExcelDna.Integration;
using KAT.Camelot.RBLe.Core;

namespace KAT.Camelot.Extensibility.Excel.AddIn.ExcelApi;

// https://excel-dna.net/docs/guides-basic/excel-programming-interfaces/excel-c-api
// https://learn.microsoft.com/en-us/office/client-developer/excel/welcome-to-the-excel-software-development-kit
static class DnaApplication
{
	public static ExcelReference? GetCaller()
	{
		var caller = XlCall.Excel( XlCall.xlfCaller );
		return caller is ExcelReference reference ? reference : null;
	}

	public static bool IsArrayFormula
	{
		get
		{
			var caller = GetCaller();
			// null if formula is being executed by xlfEvaluate
			return caller != null && caller.RowLast > caller.RowFirst;
		}
	}

	public static DnaWorkbook[] Workbooks() => ( (string[])XlCall.Excel( XlCall.xlfDocuments ) ).Select( d => new DnaWorkbook( d ) ).ToArray();

	public static string ActiveWorkbookName() => (string)XlCall.Excel( XlCall.xlfGetDocument, (int)GetDocumentType.ActiveWorkbook );
	
	public static bool ScreenUpdating
	{
		set { XlCall.Excel( XlCall.xlcEcho, value ); }
		get { return (bool)XlCall.Excel( XlCall.xlfGetWorkspace, (int)GetWorkspaceType.ScreenUpdating ); }
	}

	public static ExcelReference Selection => ( XlCall.Excel( XlCall.xlfSelection ) as ExcelReference )!;

	public static void RestoreSelection( this ExcelReference reference, Action action )
	{
		var updating = ScreenUpdating;

		try
		{
			if ( updating ) ScreenUpdating = false;

			action();

			reference.Select();
		}
		finally
		{
			if ( updating ) ScreenUpdating = true;
		}
	}

	public static ExcelReference? GetRangeFromAddress( string workbookName, string address )
	{
		try
		{
			var excelAddress = address.GetExcelAddress();
			var workbook = new DnaWorkbook( workbookName );

			var range = string.IsNullOrEmpty( excelAddress.Sheet ) || excelAddress.Sheet == workbook.Name /* global range if so */
				? workbook.RangeOrNull( excelAddress.Address )
				: null;

			if ( range != null )
			{
				return range;
			}

			if ( string.IsNullOrEmpty( excelAddress.Sheet ) )
			{
				// No named range matching
				return null;
			}

			return new DnaWorksheet( workbook.Name, excelAddress.Sheet ).RangeOrNull( excelAddress.Address );
		}
		catch ( Exception ex )
		{
			throw new ArgumentOutOfRangeException( $"Unable to convert address '{address}' into a Range object.", ex );
		}
	}
}