using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public static class ExcelExtensions
{
	private static MSExcel.Application Application => ( ExcelDnaUtil.Application as MSExcel.Application )!;
	public static MSExcel.Worksheet ActiveWorksheet( this MSExcel.Application application ) => ( application.ActiveSheet as MSExcel.Worksheet )!;

	public static void InvalidateControls( this IRibbonUI ribbon, params string[] controlIds )
	{
		foreach ( var controlId in controlIds )
		{
			ribbon.InvalidateControl( controlId );
		}
	}

	public static T? RangeOrNull<T>( this MSExcel.Workbook workbook, string name )
	{
		MSExcel.Name? namedRange = null;
		try
		{
			namedRange =
				workbook.Names.Cast<MSExcel.Name>()
					.Where( n => n.Name == name )
					.FirstOrDefault();

			var range = namedRange?.RefersToRange;

			return typeof( T ) == typeof( string )
				? (T?)range?.Text
				: (T?)range;
		}
		catch ( Exception ex )
		{
			throw new ApplicationException(
				namedRange != null
					? $"Unable to get global named range of {name}.  Currently refers to {namedRange.RefersTo}."
					: $"Unable to get global named range of {name}.", ex );
		}
	}

	public static T? RangeOrNull<T>( this MSExcel.Worksheet worksheet, string name )
	{
		var sheetName = "Unavailable";
		MSExcel.Name? namedRange = null;
		try
		{
			sheetName = worksheet.Name;
			namedRange = worksheet.Names.Cast<MSExcel.Name>()
						.Where( n => n.Name.EndsWith( "!" + name ) )
						.FirstOrDefault();

			var range = namedRange?.RefersToRange;

			return typeof( T ) == typeof( string )
				? (T?)range?.Text
				: (T?)range;
		}
		catch ( Exception ex )
		{
			throw new ApplicationException(
				namedRange != null
					? $"Unable to get global named range of {sheetName}!{name}.  Currently refers to {namedRange.RefersTo}."
					: $"Unable to get global named range of {sheetName}!{name}.", ex );
		}
	}

	public static string GetText( this MSExcel.Range range ) => ( range.Text as string )!;

	public static ExcelReference GetReference( this MSExcel.Range range )
	{
		var sheet = (ExcelReference)XlCall.Excel( XlCall.xlSheetId, range.Worksheet.Name );
		var row = range.Row - 1; // 0 based
		var column = range.Column - 1; // 0 based
		return new ExcelReference( row, row + range.Rows.Count - 1, column, column + range.Columns.Count - 1, sheet.SheetId );
	}

	/// <summary>
	/// Return MSExcel.Range from ExcelDna.ExcelReference.
	/// </summary>
	/// <remarks>
	/// Needed this because reference.GetValue() didn't preserve 'date' cells as DateTime, but rather only doubles and I'd have no way of returning
	/// DateTime values correctly.  Interop.Range.Value on the other hand preserves DateTimes.  C API had no equivalent.  So unless I determined which
	/// columns were DateTimes, and convert COM date/double to DateTime (FromOADate or something), and assuming it never changed across rows, I had no option but this method.
	/// </remarks>
	/// <param name="reference"></param>
	/// <returns></returns>
	public static MSExcel.Range GetRange( this ExcelReference reference ) => Application.Range[ reference.GetAddress() ];
}