using ExcelDna.Integration;
using KAT.Camelot.RBLe.Core;
using KAT.Camelot.RBLe.Core.Calculations;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

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

	public static ExcelReference? GetCaller()
	{
		var caller = XlCall.Excel( XlCall.xlfCaller );
		return caller is ExcelReference reference ? reference : null;
	}

	public static string ActiveWorkbookName() => (string)XlCall.Excel( XlCall.xlfGetDocument, (int)GetDocumentType.ActiveWorkbook );
	
	public static string SheetName( this ExcelReference reference ) => (string)XlCall.Excel( XlCall.xlfGetCell, (int)GetCellType.SheetRef, reference );

	public static string? GetText( this ExcelReference cell )
	{
		var value = cell.GetValue();
		return value.Equals( ExcelEmpty.Value ) ? null : (string)XlCall.Excel( XlCall.xlfGetCell, (int)GetCellType.Text, cell );
	}

	public static string? GetFormula( this ExcelReference cell )
	{
		var f = XlCall.Excel( XlCall.xlfGetCell, (int)GetCellType.Formula, cell );
		var formula = f is ExcelError check && check == ExcelError.ExcelErrorValue ? null : (string)f;
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

	public static ExcelReference Offset( this ExcelReference reference, int rows, int cols )
	{
		return new ExcelReference(
			reference.RowFirst + rows,
			reference.RowLast + rows,
			reference.ColumnFirst + cols,
			reference.ColumnLast + cols,
			reference.SheetId );
	}

	public static ExcelReference Corner( this ExcelReference reference, CornerType corner )
	{
		var row = corner == CornerType.UpperLeft || corner == CornerType.UpperRight
			? reference.RowFirst
			: reference.RowLast;

		var column = corner == CornerType.UpperLeft || corner == CornerType.LowerLeft
			? reference.ColumnFirst
			: reference.ColumnLast;

		return new ExcelReference( row, row, column, column, reference.SheetId );
	}

	public static ExcelReference GetReference( this string address ) => GetWorkbookReference( null, address );

	private static ExcelReference GetWorkbookReference( string? workbookName, string address )
	{
		return
			TryGetWorkbookReference( workbookName, address ) ?? 
			throw new ArgumentOutOfRangeException( 
				nameof( address ), 
				string.Compare( workbookName, Constants.FileNames.Helpers, true ) == 0
					? $"The address provided, {workbookName}.{address}, is not a valid address.  Please make sure you also have the '{Constants.FileNames.Helpers}' worksheet open as well."
					: $"The address provided, {workbookName}.{address}, is not a valid address."
			);
	}

	private static ExcelReference TryGetWorkbookReference( string? workbookName, string address )
	{
		var excelAddress = address.GetExcelAddress();

		var hasSheet = !string.IsNullOrEmpty( excelAddress.Sheet ) && excelAddress.Sheet != ( workbookName ?? ActiveWorkbookName() ); /* global range if so */
		if ( !hasSheet )
		{
			// Assuming if no sheet provided, they mean the 'active' sheet
			excelAddress = new()
			{
				Address = excelAddress.Address,
				Workbook = excelAddress.Workbook,
				Sheet = !string.IsNullOrEmpty( workbookName )
					? (string)XlCall.Excel( XlCall.xlfGetWorkbook, (int)GetWorkbookType.ActiveSheet, workbookName )
					: ( (string)XlCall.Excel( XlCall.xlfGetDocument, (int)GetDocumentType.ActiveSheet ) ).Split( ']' ).Last()
			};
		}

		var addressWorkbook = excelAddress.Workbook ?? workbookName;
		var workbookNameSyntax = !string.IsNullOrEmpty( addressWorkbook ) ? $"[{addressWorkbook}]" : null;

		var fullAddress = $"='{workbookNameSyntax}{excelAddress.Sheet}'!{excelAddress.Address}";
		var reference = XlCall.Excel( XlCall.xlfEvaluate, fullAddress ) as ExcelReference;

		return reference!;

		// http://www.technicana.com/xlftextref.pdf
		// Could try to get rid of all Interop objects during macro processing and attempt to use item below
		/*
		XlCall.Excel(XlCall.xlfTextref, referenceString, a1)

		referenceString is a reference in string format.

		a1 is a boolean value specifying the form the reference is in.  If a1 is TRUE, referenceString should be in A1-style.  If a1 is FALSE, referenceString must be in R1C1 format.  The default is FALSE.
		*/
	}
}