using ExcelDna.Integration;

namespace KAT.Extensibility.Excel.AddIn;

public static class ExcelApi
{
	public static ExcelReference GetCaller() => (ExcelReference)XlCall.Excel( XlCall.xlfCaller );

	public static string SheetName( this ExcelReference reference ) => (string)XlCall.Excel( XlCall.xlfGetCell, 62, reference );

	public static string? GetText( this ExcelReference cell )
	{
		var value = cell.GetValue();
		return value.Equals( ExcelEmpty.Value ) ? null : (string)XlCall.Excel( XlCall.xlfGetCell, 53, cell );
	}

	public static string? GetFormula( this ExcelReference cell )
	{
		var formula = (string)XlCall.Excel( XlCall.xlfGetCell, 6, cell );
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