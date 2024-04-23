using ExcelDna.Integration;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Extensibility.Excel.AddIn;

public static class ExcelApi
{
	enum GetCellType
	{
		Formula = 6,
		Text = 53,
		SheetRef = 62
	}

	enum GetWorkbookType
	{
		IsSaved = 24,
		ActiveSheet = 38
	}

	enum GetDocumentType
	{
		ActiveWorkbookPath = 2,
		CalculationMode = 14,
		ActiveSheet = 76, // in the form [Book1]Sheet1
		ActiveWorkbook = 88
	}

	public static ExcelReference GetCaller() => (ExcelReference)XlCall.Excel( XlCall.xlfCaller );

	public static string SheetName( this ExcelReference reference ) => (string)XlCall.Excel( XlCall.xlfGetCell, GetCellType.SheetRef, reference );

	public static string? GetText( this ExcelReference cell )
	{
		var value = cell.GetValue();
		return value.Equals( ExcelEmpty.Value ) ? null : (string)XlCall.Excel( XlCall.xlfGetCell, GetCellType.Text, cell );
	}

	public static string? GetFormula( this ExcelReference cell )
	{
		var formula = (string)XlCall.Excel( XlCall.xlfGetCell, GetCellType.Formula, cell );
		return !string.IsNullOrEmpty( formula ) ? formula : null;
	}

	public static string GetAddress( this ExcelReference? reference )
	{
		try
		{
			var address = (string)XlCall.Excel( XlCall.xlfReftext, reference, true /* true - A1, false - R1C1 */ );
			return address;
		}
		catch ( Exception ex )
		{
			throw new ApplicationException( $"GetAddress failed.  reference.RowFirst:{reference?.RowFirst}, reference.RowLast:{reference?.RowLast}, reference.ColumnFirst:{reference?.ColumnFirst}, reference.ColumnLast:{reference?.ColumnLast}", ex );
		}
	}
}