using ExcelDna.Integration;

namespace KAT.Camelot.Extensibility.Excel.AddIn.Functions;

#pragma warning disable IDE0060 // Remove unused parameter

public static class Macro
{
	[ExcelFunction( Category = "RBLe Macro", Description = "Returns variable value set during macro processing.  If not during macro processing, this returns null." )]
	public static object BTRGetMacroVariable(
		[ExcelArgument( "The name of the variable stored via SetVariable action during RBLe Macro processing you wish to retrieve.  This is a placeholder function returning 'empty' when used in Excel, but during ProcessWorkbook execution it evaluates properly." )]
		string name
	) => ExcelEmpty.Value;

	// See ServiceHeleprs.GetMacroSafeFormula and Core.BTRCellAddress for more details on hidden function reason
	[ExcelFunction( IsHidden = true, IsVolatile = true, IsMacroType = true )]
	public static string BTRCellAddressMacro(
		[ExcelArgument( AllowReference = true )]
		object range,
		int rowOffset = 0, int columnOffset = 0, int additionalRows = 0, int additionalColumns = 0 
	) => BTRCellAddress( range, rowOffset, columnOffset, additionalRows, additionalColumns );

	[ExcelFunction( Category = "RBLe Macro", Description = "Returns the address (similar to Cell('address', cell) format) of a named range to be used in RBLe Macro processing.", IsVolatile = true, IsMacroType = true )]
	public static string BTRCellAddress(
		[ExcelArgument( AllowReference = true, Description = "The cell or range used as starting point of address." )]
		object range,
		[ExcelArgument( "The number of rows (positive, negative, or 0 (zero)) by which the range is to be offset. Positive values are offset downward, and negative values are offset upward. The default value is 0." )]
		int rowOffset = 0,
		[ExcelArgument( "The number of columns (positive, negative, or 0 (zero)) by which the range is to be offset. Positive values are offset to the right, and negative values are offset to the left. The default value is 0." )]
		int columnOffset = 0,
		[ExcelArgument( "The number of rows (positive, negative, or 0 (zero)) by which the range is to be resized. Positive values grow downward, and negative values shrink upward. The default value is 0." )]
		int additionalRows = 0,
		[ExcelArgument( "The number of columns (positive, negative, or 0 (zero)) by which the range is to be resized. Positive values grow to the right, and negative values shrink to the left. The default value is 0." )]
		int additionalColumns = 0 
	)
	{
		try
		{
			var reference = ( range as ExcelReference )!;

			// var nameBeforeOffset = (string)XlCall.Excel( XlCall.xlSheetNm, range );

			// Pending question: https://groups.google.com/forum/#!topic/exceldna/YKyYoUjVOC4
			// Note, the problem described in post only happens if/when following:
			// 1. Run ProcessWorkbook when the active workbook is Calc (when RBLMacro, it works fine)
			// 2. After instruction $A$5 happens, which is doing CopyValue "0" to 'Calc'!$AE$6 (which is iSaveRateVary), then the next call to BTRCellAddress(iSaveRateVary) fails (if I skip assignment, it works)

			// range = XlCall.Excel( XlCall.xlfOffset, range, rowOffset, columnOffset ) as ExcelReference;
			// So instead of the XlCall, I made a new ExcelReference by hand...
			reference = new ExcelReference( reference.RowFirst + rowOffset, reference.RowLast + rowOffset, reference.ColumnFirst + columnOffset, reference.ColumnLast + columnOffset, reference.SheetId );

			var rowCount = reference.RowLast - reference.RowFirst + 1;
			var columnCount = reference.ColumnLast - reference.ColumnFirst + 1;

			var rowFirst = ( rowCount + additionalRows - 1 ) >= 0
				? reference.RowFirst
				: reference.RowFirst + rowCount + additionalRows - 1;
			var rowLast = ( rowCount + additionalRows - 1 ) >= 0
				? reference.RowFirst + rowCount + additionalRows - 1
				: reference.RowFirst;

			var colFirst = ( columnCount + additionalColumns - 1 ) >= 0
				? reference.ColumnFirst
				: reference.ColumnFirst + columnCount + additionalColumns - 1;
			var colLast = ( columnCount + additionalColumns - 1 ) >= 0
				? reference.ColumnFirst + columnCount + additionalColumns - 1
				: reference.ColumnFirst;

			// System.Diagnostics.Trace.WriteLine( $"Start: {XlCall.Excel( XlCall.xlfReftext, reference, true )}, RowOffset: {rowOffset}, ColOffset: {columnOffset}, RowFirst: {rowFirst}, RowLast: {rowLast}, ColFirst: {colFirst}, ColLast: {colLast}" );

			reference = new ExcelReference( rowFirst, rowLast, colFirst, colLast, reference.SheetId );

			var address = (string)XlCall.Excel( XlCall.xlfReftext, reference, true );
			var workbookMarker = -1;
			if ( ( workbookMarker = address.IndexOf( "]" ) ) > -1 )
			{
				// Remove ' in case workbook name had space so address was '[name with space]Sheet'!A1
				address = address[ ( workbookMarker + 1 ).. ].Replace( "'", "" );
			}

			// System.Diagnostics.Trace.WriteLine( $"Start: {start}, End: {address}" );

			/*
			var worksheetName = (string)XlCall.Excel( XlCall.xlSheetNm, reference ); //[Workbook]Sheet

			// This was getting like the active workbook or something, anyway, if multi workbooks open, this didn't necessarily return name of workbook
			// of the namedRange...so had to use the xlSheetNm function instead.
			// var workbookName = (string)XlCall.Excel( XlCall.xlfGetWorkbook, 16 );

			var bookPrefix = worksheetName.Substring( 0, worksheetName.IndexOf( "]" ) + 1 ); // string.Format( "[{0}]", workbookName );

			if ( !address.Contains( bookPrefix ) )
			{
				System.Diagnostics.Debugger.Break();
			}

			return address.Replace( bookPrefix, "" );
			*/

			return address;
		}
		catch ( Exception ex )
		{
			System.Diagnostics.Trace.WriteLine( "BTRCellAddress: " + ex.Message );
			throw;
		}
	}

	/* Can't safely support these as they would be called during SSG's internal 'recalculation' process and using excel object model APIs is not 'safe'
	[ExcelFunction( IsVolatile = true, IsMacroType = true, Category = "Information", Description = "Returns number of rows between current cell and 'EndDown'." )]
	public static int BTRRowsToBottom( [ExcelArgument( AllowReference = true )] object cell )
	{
		var currentCell = cell as ExcelReference;
		var range = currentCell.GetRange();
		var row = range.get_End( MSExcel.XlDirection.xlDown ).Row;
		/*
			+-----------------+-----------+--------------+---------------------+
			|                 | Max. Rows | Max. Columns | Max. Cols by letter |
			+-----------------+-----------+--------------+---------------------+
			| Excel 365*      | 1,048,576 | 16,384       | XFD                 |
			| Excel 2013      | 1,048,576 | 16,384       | XFD                 |
			| Excel 2010      | 1,048,576 | 16,384       | XFD                 |
			| Excel 2007      | 1,048,576 | 16,384       | XFD                 |
			| Excel 2003      | 65,536    | 256          | IV                  |
			| Excel 2002 (XP) | 65,536    | 256          | IV                  |
			| Excel 2000      | 65,536    | 256          | IV                  |
			| Excel 97        | 65,536    | 256          | IV                  |
			| Excel 95        | 16,384    | 256          | IV                  |
			| Excel 5         | 16,384    | 256          | IV                  |
			+-----------------+-----------+--------------+---------------------+
		 * /
		return row == 1048576
			? 0
			: Math.Max( 0, row - currentCell.RowFirst - 1 );
	}

	[ExcelFunction( IsVolatile = true, IsMacroType = true, Category = "Information", Description = "Returns number of columns between current cell and 'EndRight'." )]
	public static int BTRColumnsToRight( [ExcelArgument( AllowReference = true )] object cell )
	{
		var currentCell = cell as ExcelReference;
		var range = currentCell.GetRange();
		var col = range.get_End( MSExcel.XlDirection.xlToRight ).Column;
		return col == 16384
			? 0
			: Math.Max( 0, col - currentCell.ColumnFirst - 1 );
	}
	*/
}
