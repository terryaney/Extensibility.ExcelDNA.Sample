using ExcelDna.Integration;
using KAT.Camelot.Domain.Extensions;
using KAT.Camelot.RBLe.Core.Calculations.Functions;

namespace KAT.Camelot.Extensibility.Excel.AddIn.Functions;

public static class Logic
{
	[ExcelFunction( Category = "Validation", Description = "Returns 1 if all parameters provided evaluate to not 'false' (#error, 0, FALSE, '') or 0 if any parameter provided is 'false'." )]
	public static int BTROnAll(
		[ExcelArgument( "The parameters to evaluate falsy conditions." )]
		params object[] parameters
	)
	{
		return Validation.LogicAnd(
			v => v is ExcelError,
			ExcelEmpty.Value,
			parameters
		);
	}

	[ExcelFunction( Category = "Validation", Description = "Returns 1 if any parameters provided evaluate to not 'false' (#error, 0, FALSE, '') or 0 if all parameters provided are 'false'." )]
	public static int BTROnAny(
		[ExcelArgument( "The parameters to evaluate falsy conditions." )]
		params object[] parameters
	)
	{
		return Validation.LogicOr(
			v => v is ExcelError,
			ExcelEmpty.Value,
			parameters
		);
	}

	[ExcelFunction( Category = "Validation", Description = "Returns 1 if all keys provided evaluate to not 'false' (#error, 0, FALSE) or 0 if any key provided is 'false'.  Each 'key' is a lookup into provided table, and the 'value' column is evalated for 'false'." )]
	public static int BTRLogicAll(
		[ExcelArgument( "Comma delimitted list of key values to use to find the rows in the table parameter." )]
		string keys,
		[ExcelArgument( "Table containing the columns holding keys and values to find and do a falsy check (#error, 0, FALSE)." )]
		object[,] table,
		[ExcelArgument( "The column (1..table columns) number containing the value to compare to ensure not falsy.  Default is the last column of 'table' parameter." )]
		object? valueColumn = null,
		[ExcelArgument( "Whether or not a case sensitive search is performed. Default is false." )]
		object? caseSensitive = null
	)
	{
		if ( string.IsNullOrWhiteSpace( keys ) ) return 0;

		var tableColumns = table.GetColumns();
		var totalColumns = tableColumns.Length;

		if ( totalColumns == 0 ) return 0;

		// If no valueColumnArg passed, default to -1
		var valueColumnArg = valueColumn.Check( nameof( valueColumn ), totalColumns );

		if ( valueColumnArg <= 1 || valueColumnArg > totalColumns )
		{
			throw new ArgumentOutOfRangeException( nameof( valueColumn ), $"{nameof( valueColumn )} must be greater than 1 and less than {totalColumns}." );
		}

		return Validation.LogicAll(
			keys,
			tableColumns,
			valueColumnArg,
			v => v is ExcelError,
			ExcelEmpty.Value,
			caseSensitive.Check( nameof( caseSensitive ), false )
		);
	}

	[ExcelFunction( Category = "Validation", Description = "Returns 1 if any key provided evaluates to not 'false' (#error, 0, FALSE) or 0 if all keys provided are 'false'.  Each 'key' is a lookup into provided table, and the 'value' column is evalated for 'false'." )]
	public static int BTRLogicAny(
		[ExcelArgument( "Comma delimitted list of key values to use to find the rows in the table parameter." )]
		string keys,
		[ExcelArgument( "Table containing the columns holding keys and values to find and do a falsy check (#error, 0, FALSE)." )]
		object[,] table,
		[ExcelArgument( "The column (1..table columns) number containing the value to compare to ensure not falsy.  Default is the last column of 'table' parameter." )]
		object? valueColumn = null,
		[ExcelArgument( "Whether or not a case sensitive search is performed. Default is false." )]
		object? caseSensitive = null
	)
	{
		if ( string.IsNullOrWhiteSpace( keys ) ) return 0;

		var tableColumns = table.GetColumns();
		var totalColumns = tableColumns.Length;

		if ( totalColumns == 0 ) return 0;

		// If no valueColumnArg passed, default to -1
		var valueColumnArg = valueColumn.Check( nameof( valueColumn ), totalColumns );

		if ( valueColumnArg <= 1 || valueColumnArg > totalColumns )
		{
			throw new ArgumentOutOfRangeException( nameof( valueColumn ), $"{nameof( valueColumn )} must be greater than 1 and less than {totalColumns}." );
		}

		return Validation.LogicAny(
			keys,
			tableColumns,
			valueColumnArg,
			v => v is ExcelError,
			ExcelEmpty.Value,
			caseSensitive.Check( nameof( caseSensitive ), false )
		);
	}

	[ExcelFunction( Category = "Validation", Description = "Given param array of {condition1}, {expression1}, {condition2}, {expression2}, .., {conditionN, {expressionN}, returns expression value for the first condition that evalates to true.  If all conditions are false, returns #NA." )]
	public static object? BTRLogicIfElse(
		[ExcelArgument( "List of 'paired' parameters.  Each pair is condition, expression.  If condition is true, return expression." )]
		params object[] expressions
	)
	{
		if ( expressions.Length % 2 != 0 )
		{
			throw new ArgumentOutOfRangeException( nameof( expressions ), $"{nameof( expressions )} should have paired values of {{condition1}}, {{expression1}}, ...N." );
		}

		return Validation.LogicIfElse( 
			v => v is ExcelError, 
			ExcelEmpty.Value, 
			ExcelError.ExcelErrorNA, 
			expressions 
		);
	}

	[ExcelFunction( Category = "Validation", Description = "Returns 1 if 'value' is present anywhere in the 'table' parameter." )]
	public static int BTRLogicIn(
		[ExcelArgument( "The value to search for." )]
		object value,
		[ExcelArgument( "Table containing values to search." )]
		object[,] table,
		[ExcelArgument( "Whether or not a case sensitive search is performed. Default is false." )]
		object? caseSensitive = null
	)
	{
		var caseSensitiveArg = caseSensitive.Check( nameof( caseSensitive ), false );

		if ( typeof( object[,] ) == value?.GetType() )
		{
			return table.ContainsAny( (object[,])value, caseSensitiveArg ) ? 1 : 0;
		}
		else
		{
			var stringValue = (string?)value;
			return !string.IsNullOrWhiteSpace( stringValue ) && table.Contains( stringValue, caseSensitiveArg )
				? 1
				: 0;
		}
	}
}