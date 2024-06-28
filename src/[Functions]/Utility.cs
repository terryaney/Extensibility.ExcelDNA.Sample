using ExcelDna.Integration;
using KAT.Camelot.Domain.Extensions;
using KAT.Camelot.Extensibility.Excel.AddIn.ExcelApi;
using KAT.Camelot.RBLe.Core.Calculations.Functions;

namespace KAT.Camelot.Extensibility.Excel.AddIn.Functions;

public static class DnaUtility
{
	[ExcelFunction( Category = "General Utility Helpers", Description = "Returns whether a specified substring occurs within a string (optionally case sensitive)." )]
	public static bool BTRContains(
		[ExcelArgument( "The text to search." )]
		string text,
		[ExcelArgument( "The text to search for." )]
		string find,
		[KatExcelArgument(
			Description = "Optional.  Whether or not the search is case insensitive or not.  True is the default.",
			Type = typeof( bool ),
			Default = "true"
		)]
		object? caseInsensitive = null
	)
	{
		var caseInsensitiveArg = caseInsensitive.Check( nameof( caseInsensitive ), true );

		if ( string.IsNullOrEmpty( find ) )
		{
			return false;
		}

		return caseInsensitiveArg
			? text.IndexOf( find, StringComparison.InvariantCultureIgnoreCase ) > -1
			: text.Contains( find );
	}

	[ExcelFunction( 
		Category = "General Utility Helpers", 
		Description = "Returns a filtered table given a list of filter columns and values to compare." 
	)]
	public static object[,] BTRFilter(
		[ExcelArgument( "Range of cells to filter." )]
		object[,] table,
		[ExcelArgument( "Comma delimitted list of column names that should be included in returned range." )]
		string returnColumns,
		[KatExcelArgument(
			Description = "Optional.  Whether or not to include column headers in result table. Default is true.",
			Type = typeof( bool ),
			Default = "true"
		)]
		object? includeHeadersInResults = null,
		[ExcelArgument( "Paired expressions.  First item is column name to search, second item is value to filter." )]
		params object[] expressions
	)
	{
		var returnColumnNames = returnColumns.Split( ',' ).Select( c => c.Trim() ).Where( c => !string.IsNullOrEmpty( c ) ).ToArray();

		if ( returnColumnNames.Length == 0 )
		{
			return new object[,] { { ExcelError.ExcelErrorValue } };
		}

		if ( expressions.Length % 2 != 0 )
		{
			return new object[,] { { ExcelError.ExcelErrorValue } };
			// throw new ArgumentOutOfRangeException( nameof( expressions ), $"{nameof( expressions ) } should have paired values of {{columnNumber}}, {{filterValue}}, ...N." );
		}

		var filterResult = Utility.Filter( table, includeHeadersInResults.Check( nameof( includeHeadersInResults ), true ), returnColumnNames, expressions, ExcelError.ExcelErrorNA );

		var isArrayFunction = DnaApplication.IsArrayFormula;

		if ( isArrayFunction && filterResult.GetUpperBound( 0 ) == 0 && !object.Equals( filterResult[ 0, 0 ], ExcelError.ExcelErrorNA ) )
		{
			var output = new object[ filterResult.GetUpperBound( 0 ) + 2, filterResult.GetUpperBound( 1 ) + 1 ];

			for ( var row = 0; row <= filterResult.GetUpperBound( 0 ); row++ )
			{
				for ( var c = 0; c <= filterResult.GetUpperBound( 1 ); c++ )
				{
					output[ row, c ] = filterResult[ row, c ];
				}
			}

			for ( var c = 0; c <= filterResult.GetUpperBound( 1 ); c++ )
			{
				output[ filterResult.GetUpperBound( 0 ) + 1, c ] = ExcelError.ExcelErrorNA;
			}

			return output;
		}

		return filterResult;
	}

	[ExcelFunction( Category = "General Utility Helpers", Description = "Returns a spine-cased xDS name from PascalCase or iInputName input string." )]
	public static string BTRToxDSName( string value ) => value.ToxDSName()!;

	[ExcelFunction( Category = "General Utility Helpers", Description = "Returns an iInputName from a spine-case string." )]
	public static string BTRToInputName( string value ) => value.ToInputName();

	[ExcelFunction( Category = "General Utility Helpers", Description = "Returns a camelCase string from a spine-case string." )]
	public static string BTRToCamelCase( string source ) => source.ToCamelCase();

	[ExcelFunction( Category = "General Utility Helpers", Description = "Returns a EBCDIC, extended binary-coded decimal interchange code, encoded string." )]
	public static string BTREbcdicText( string value, int padLength ) => Utility.GetEbcdicText( value, padLength );

	[KatExcelFunction( 
		Category = "General Utility Helpers", 
		Description = "Returns the value (or fallback is not found) from columnToReturn column in a table range (including column headers).  Similar to VLOOKUP, but finds exact match, searches specified column versus always the first, and allows for a default.",
		Summary = "Returns the value (or fallback is not found) from columnToReturn column in a table range (including column headers).  Similar to VLOOKUP but always finds exact match, uses column names instead of numbers (prevents issues when columns are inserted or removed), can search specified column instead of always the first, and provides ability to give a default."
	)]
	public static object BTRLookup(
		[ExcelArgument( "Value to search for." )]
		object[,] values,
		[ExcelArgument( "Range of cells to search (first row must be column headers)." )]
		object[,] table,
		[KatExcelArgument(
			Description = "Optional. Column name containing return value. Last column is default",
			Type = typeof( string )
		)]
		object? columnToReturn = null,
		[KatExcelArgument(
			Description = "Optional. Column name to search.  First column is default.",
			Type = typeof( string )
		)]
		object? columnToSearch = null,
		[KatExcelArgument(
			Description = "Optional.  Value to return if a match is not found.  #N/A is the default.",
			Type = typeof( object ),
			Default = "#N/A"
		)]
		object? fallback = null,
		[KatExcelArgument(
			Description = "Optional.  Whether or not search is case sensitive. false is the default.",
			Type = typeof( bool ),
			Default = "false"
		)]
		object? caseSensitive = null
	)
	{
		var isArrayFunction = DnaApplication.IsArrayFormula;

		var returnValues = Utility.LookupValues(
			table,
			columnToSearch.Check<string?>( nameof( columnToSearch ), null ),
			columnToReturn.Check<string?>( nameof( columnToReturn ), null ),
			values.ToArray(),
			fallback.Check<object?>( nameof( fallback ), null ) ?? ExcelError.ExcelErrorNA,
			caseSensitive.Check( nameof( caseSensitive ), false ),
			ExcelEmpty.Value
		);

		return isArrayFunction ? returnValues : returnValues[ 0, 0 ]!;
	}

	[KatExcelFunction( 
		Category = "General Utility Helpers", 
		Description = "Concatenates all values from columnToReturn in a table with separator between values.  Similar to VLOOKUP but finds exact match, uses column names versus numbers, searches column versus always the first, and provides ability far a default.",
		Summary = "Concatenates all values (or fallback is not found) from columnToReturn column in a table range (including column headers), using the specified separator between elements.  Similar to VLOOKUP but always does exact match, uses column names instead of numbers (prevents issues when columns are inserted or removed), can search specified column instead of always the first, and provides ability to give a default."
	)]
	public static string BTRJoinLookup(
		[ExcelArgument( "Comma delimitted list of values to search for." )]
		string values,
		[ExcelArgument( "Range of cells to search (first row must be column headers)." )]
		object[,] table,
		[KatExcelArgument(
			Description = "Optional. Column name containing return value. Last column is default",
			Type = typeof( string )
		)]
		object? columnToReturn = null,
		[KatExcelArgument(
			Description = "Optional. Column name to search.  First column is default.",
			Type = typeof( string )
		)]
		object? columnToSearch = null,
		[KatExcelArgument(
			Description = "Optional.  Value to return if a match is not found.  #N/A is the default.",
			Type = typeof( object ),
			Default = "#N/A"
		)]
		object? fallback = null,
		[KatExcelArgument(
			Description = "Optional.  Whether or not search is case sensitive. false is the default.",
			Type = typeof( bool ),
			Default = "false"
		)]
		object? caseSensitive = null,
		[KatExcelArgument(
			Description = "Optional. The string to use as a separator.  The separator is included in the return string only if values has more than one element. ', ' is the default.",
			Type = typeof( string ),
			Default = ", "
		)]
		object? separator = null
	) =>
		Utility.JoinLookup(
			table,
			columnToSearch.Check<string?>( nameof( columnToSearch ), null ),
			columnToReturn.Check<string?>( nameof( columnToReturn ), null ),
			values.Split( ',' ).Select( v => v.Trim() ).ToArray(),
			fallback ?? ExcelError.ExcelErrorNA,
			caseSensitive.Check( nameof( caseSensitive ), false ),
			ExcelEmpty.Value,
			separator.Check<string?>( nameof( separator ), null ) ?? ", "
		);

	[ExcelFunction( Category = "General Utility Helpers", Description = "Evaluates an Excel formula passed in.  If the formula contains one or more {K.id} tokens, each is parsed and the 'id' is used in BTRLookup to find a value to substitute into the formula", IsMacroType = true )]
	public static object BTREvaluate(
		[ExcelArgument( "Formula to evaluate with optional {K.id} tokens." )]
		string formula,
		[ExcelArgument( "Range of cells to search (first row must be column headers)." )]
		object[,] table,
		[KatExcelArgument(
			Description = "Optional. Column name containing return value. Last column is default",
			Type = typeof( string )
		)]
		object? columnToReturn = null,
		[KatExcelArgument(
			Description = "Optional. Column name to search.  First column is default.",
			Type = typeof( string )
		)]
		object? columnToSearch = null,
		[KatExcelArgument(
			Description = "Optional.  Whether or not search is case sensitive. false is the default.",
			Type = typeof( bool ),
			Default = "false"
		)]
		object? caseSensitive = null
	)
	{
		if ( table[ 0, 0 ] != ExcelMissing.Value )
		{
			formula = Utility.GetEvaluateFormula(
				formula,
				table,
				columnToSearch.Check<string?>( nameof( columnToSearch ), null ),
				columnToReturn.Check<string?>( nameof( columnToReturn ), null ),
				ExcelError.ExcelErrorNA,
				caseSensitive.Check( nameof( caseSensitive ), false ),
				ExcelEmpty.Value );
		}

		var value = XlCall.Excel( XlCall.xlfEvaluate, formula );
		return value;
	}
}