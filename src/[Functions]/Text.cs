using ExcelDna.Integration;
using KAT.Camelot.Domain.Extensions;
using KAT.Camelot.Extensibility.Excel.AddIn.ExcelApi;
using KAT.Camelot.RBLe.Core.Calculations.Functions;

namespace KAT.Camelot.Extensibility.Excel.AddIn.Functions;

public static class Text
{
	/// <summary>
	/// Formats a numeric value to a string representation using the specified format and culture-specific format information.
	/// </summary>
	/// <param name="value">The numeric value to apply formatting to.</param>
	/// <param name="format">A numeric format string.</param>
	/// <param name="culture">An string that supplies culture-specific formatting information.  The default value is <c>en-US</c>.</param>
	/// <returns>The string representation of the value of this instance as specified by <paramref name="format"/> and <paramref name="culture"/>.</returns>
	/// <remarks>The BTRNumberFormat method is similar to Excel's Format() function with the exception that BTRNumberFormat can dynamically format a number based on <paramref name="culture"/> using the same <paramref name="format"/> string.</remarks>
	/// <seealso href="http://msdn.microsoft.com/en-us/library/dwhawy9k(v=vs.110).aspx">Standard Numeric Format Strings</seealso>
	/// <seealso href="http://msdn.microsoft.com/en-us/library/0c899ak8(v=vs.110).aspx">Custom Numeric Format Strings</seealso>
	/// <example>
	/// This sample shows how to format a numeric value to currency format with a single format string but changes based on culture.
	/// <code>
	/// // Assume this comes from the iCurrentCulture input.
	/// string culture = "en-US";
	/// // Assume this comes from a calculated result.
	/// double value = 10.5;
	/// // currencyValue would have "$10.50" for a value.
	/// string currencyValue = BTRNumberFormat( value, "c", culture );
	/// // If culture was French...
	/// culture = "fr-FR";
	/// // currencyValue would have "10,50 €" for a value.
	/// currencyValue = BTRNumberFormat( value, "c", culture );
	/// </code>
	/// </example>
	[ExcelFunction( Category = "Formatting", Description = "Formats a numeric value to a string representation using the specified format and culture-specific format information." )]
	public static string BTRNumberFormat(
		[ExcelArgument( "The number to apply a format to." )]
		double value,
		[ExcelArgument( "The C# string format to apply.  View the function's help for more detail on possible values." )]
		string format,
		[ExcelArgument( "Optional.  The culture name in the format languagecode2-country/regioncode2.  See 'National Language Support (NLS) API Reference' for available names." )]
		object? culture = null
	)
	{
		var cultureArg = culture.Check( nameof( culture ), "en-US" );
		return Utility.LocaleFormat( value, format, cultureArg );
	}

	/// <summary>
	/// Formats a Date value to a string representation using the specified format and culture-specific format information.
	/// </summary>
	/// <param name="value">The Date value to apply formatting to.</param>
	/// <param name="format">A numeric format string.</param>
	/// <param name="culture">An string that supplies culture-specific formatting information.  The default value is <c>en-US</c>.</param>
	/// <returns>The string representation of the value of this instance as specified by <paramref name="format"/> and <paramref name="culture"/>.</returns>
	/// <remarks>The BTRDateFormat method is similar to Excel's Format() function with the exception that BTRDateFormat can dynamically format a date based on <paramref name="culture"/> using the same <paramref name="format"/> string.</remarks>
	/// <seealso href="http://msdn.microsoft.com/en-us/library/az4se3k1(v=vs.110).aspx">Standard Date and Time Format Strings</seealso>
	/// <seealso href="http://msdn.microsoft.com/en-us/library/8kb3ddd4(v=vs.110).aspx">Custom Date and Time Format Strings</seealso>
	/// <example>
	/// This sample shows how to format a date value to 'short date' format with a single format string, but changes based on culture.
	/// <code>
	/// // Assume this comes from the iCurrentCulture input.
	/// string culture = "en-US";
	/// // Assume this comes from a calculated result.
	/// DateTime value = new DateTime( 1973, 5, 9 );
	/// // dateValue would have "5/9/1973" for a value.
	/// string dateValue = BTRDateFormat( value, "d", culture );
	/// // If culture was French...
	/// culture = "fr-FR";
	/// // currencyValue would have "09/05/1973" for a value.
	/// dateValue = BTRDateFormat( value, "d", culture );
	/// </code>
	/// </example>
	[ExcelFunction( Category = "Formatting", Description = "Formats a Date value to a string representation using the specified format and culture-specific format information." )]
	public static string BTRDateFormat(
		[ExcelArgument( "The date to apply a format to." )]
		object value,
		[ExcelArgument( "The C# string format to apply.  View the function's help for more detail on possible values." )]
		string format,
		[ExcelArgument( "Optional.  The culture name in the format languagecode2-country/regioncode2.  See 'National Language Support (NLS) API Reference' for available names." )]
		object? culture = null 
	)
	{
		var cultureArg = culture.Check( nameof( culture ), "en-US" );
		// return Functions.Utility.LocaleFormat( Functions.Validation.ParseDate( value ), format, culture );

		var date = value switch
		{
			double d => DateTime.FromOADate( d ),
			string s => DateTime.Parse( s ),
			DateTime dt => dt,
			_ => throw new ArgumentOutOfRangeException( nameof( value ) )
		};

		return Utility.LocaleFormat( 
			date < new DateTime( 1900, 3, 1 ) ? date.AddDays( 1 ) : date, 
			format, 
			cultureArg 
		);
	}

	/// <summary>
	/// Builds a string using the template with zero based subsitution tokens (i.e. {0}, {1}, ...) and substitutes them for the supplied parameters.
	/// </summary>
	/// <param name="template">The string template to use in the builder with zero based subsitution tokens (i.e. {0}, {1}, ...).</param>
	/// <param name="parameters">The parameters to substitute into the string template.</param>
	/// <returns>The string after subsituting all <paramref name="parameters"/> into the <paramref name="template"/>.</returns>
	/// <remarks>The BTRStringBuilder method is similar to C#'s string.Format() function.</remarks>
	/// <seealso href="http://msdn.microsoft.com/en-us/library/dwhawy9k(v=vs.110).aspx">Standard Numeric Format Strings</seealso>
	/// <seealso href="http://msdn.microsoft.com/en-us/library/0c899ak8(v=vs.110).aspx">Custom Numeric Format Strings</seealso>
	[ExcelFunction( Category = "Formatting", Description = "Builds a string using the template with zero based subsitution tokens (i.e. {0}, {1}, ...) and substitutes them for the supplied parameters." )]
	public static string BTRStringBuilder(
		[ExcelArgument( "The string template to use in the builder with zero based subsitution tokens (i.e. {0}, {1}, ...)." )]
		string template,
		[ExcelArgument( "The parameters to substitute into the string template." )]
		params object[] parameters
	) => Utility.StringFormat( null, template, parameters, ExcelEmpty.Value, typeof( ExcelError ), DateTime.FromOADate );

	/// <summary>
	/// Builds a string using the template with zero based subsitution tokens (i.e. {0}, {1}, ...) and substitutes them for the supplied parameters.
	/// </summary>
	/// <param name="placeHolders">A space delimitted open and closing placeholder to use for token matching (i.e. {{ }}, << >>, or < >).</param>
	/// <param name="template">The string template to use in the builder with zero based subsitution tokens (i.e. {0}, {1}, ...).</param>
	/// <param name="parameters">The parameters to substitute into the string template.</param>
	/// <returns>The string after subsituting all <paramref name="parameters"/> into the <paramref name="template"/>.</returns>
	/// <remarks>The BTRStringBuilder method is similar to C#'s string.Format() function.</remarks>
	/// <seealso href="http://msdn.microsoft.com/en-us/library/dwhawy9k(v=vs.110).aspx">Standard Numeric Format Strings</seealso>
	/// <seealso href="http://msdn.microsoft.com/en-us/library/0c899ak8(v=vs.110).aspx">Custom Numeric Format Strings</seealso>
	[ExcelFunction( Category = "Formatting", Description = "Builds a string using the template with zero based subsitution tokens (i.e. {0}, {1}, ...) and substitutes them for the supplied parameters." )]
	public static string BTRStringBuilderWithPlaceholder(
		[ExcelArgument( "A space delimitted open and closing placeholder to use for token matching (i.e. {{ }}, << >>, or < >)." )]
		string placeHolders,
		[ExcelArgument( "The string template to use in the builder with zero based subsitution tokens (i.e. {0}, {1}, ...)." )]
		string template,
		[ExcelArgument( "The parameters to substitute into the string template." )]
		params object[] parameters
	) => Utility.StringFormat( placeHolders.Split( ' ' ), template, parameters, ExcelEmpty.Value, typeof( ExcelError ), DateTime.FromOADate );
	
	/// <summary>
	/// Joins a range of text strings into one string using seperator.
	/// </summary>
	/// <param name="values">The string values to join together</param>
	/// <param name="seperator">Optional seperator to use between the strings (empty string if omitted).</param>
	/// <returns>The string after joining all <paramref name="values"/> with the <paramref name="seperator"/>.</returns>
	[ExcelFunction( Category = "Formatting", Description = "Joins a range of text strings into one string using seperator." )]
	public static string BTRJoin(
		[ExcelArgument( "The string values to join with last parameter being the separator." )]
		params object[] argsAndSeparator
	)
	{
		if ( argsAndSeparator.Length < 2 )
		{
			throw new ArgumentOutOfRangeException( nameof( argsAndSeparator ), "You must pass in at least two parameters with the last parameter being a string representing the separator used in the join." );
		}

		IEnumerable<object> arguments = Array.Empty<object>();

		foreach ( var argument in argsAndSeparator.Take( argsAndSeparator.Length - 1 ) )
		{
			arguments = arguments.Concat( ( argument as object[,] )?.ToArray() ?? new object[] { argument } );
		}

		var seperatorArg = argsAndSeparator.Last() as string;

		return string.Join( seperatorArg, arguments.Where( o => o != ExcelEmpty.Value && !string.IsNullOrEmpty( o?.ToString() ) ) );
	}

	[ExcelFunction( Category = "Formatting", Description = "Concatenates a list or range of text strings using a delimiter. Polyfill for non-supported Excel TEXTJOIN function." )]
	public static string BTRTextJoin(
		[ExcelArgument( "Optional. The separator to use between values. Default is empty string." )]
		object? seperator = null,
		[ExcelArgument( "Optional. Whether or not to ignore empty cells.  Default is true" )]
		object? ignoreEmptyCells = null,
		[ExcelArgument( "1 to 252 text strings or ranges to be joined." )]
		params object[] ranges
	)
	{
		var seperatorArg = seperator.Check( nameof( seperator ), "" );
		var ignoreEmptyCellsArg = ignoreEmptyCells.Check( nameof( ignoreEmptyCells ), true );

		var values = new List<object>();
		
		foreach ( var range in ranges )
		{
			if ( range is object[,] array )
			{
				values.AddRange( array.ToArray().Where( o => !ignoreEmptyCellsArg || ( o != ExcelEmpty.Value && !string.IsNullOrEmpty( o?.ToString() ) ) ) );
			}
			else if ( !ignoreEmptyCellsArg || ( range != ExcelEmpty.Value && !string.IsNullOrEmpty( range?.ToString() ) ) )
			{
				values.Add( range );
			}
		}
		return string.Join( seperatorArg, values );
	}

	[ExcelFunction( Category = "Formatting", Description = "Returns a list of unique values from a given input range.", IsMacroType = true )]
	public static object[,] BTRUnique(
		[ExcelArgument( "Range of cells to find unique values contained." )]
		object[,] values,
		object? matchInputOutputSize = null,
		[ExcelArgument( "Optional. Whether or not to ignore empty cells.  Default is true" )]
		object? ignoreEmptyCells = null,
		[ExcelArgument( "1 to 252 ranges to merge." )]
		params object[] ranges
	)
	{
		var isArrayFunction = DnaApplication.IsArrayFormula;

		var list = new HashSet<object>();
		var ignoreEmptyCellsArg = ignoreEmptyCells.Check( nameof( ignoreEmptyCells ), true );
		var matchInputOutputSizeArg = ranges.Length == 0 && matchInputOutputSize.Check( nameof( matchInputOutputSize ), true );

		void processArray( object[] array )
		{
			foreach ( var v in array )
			{
				var validValue =
					!ignoreEmptyCellsArg ||
					( v != ExcelEmpty.Value && !string.IsNullOrEmpty( v?.ToString() ) );

				if ( validValue && !list.Contains( v! ) )
				{
					list.Add( v != ExcelEmpty.Value ? v! : "" );
				}
			}
		}

		processArray( values.ToArray() );

		foreach ( var range in ranges )
		{
			if ( range is object[,] array )
			{
				processArray( array.ToArray() );
			}
			else
			{
				processArray( new[] { range } );
			}
		}

		var vals = list.ToArray();
		var arrayLength = matchInputOutputSizeArg ? values.Length : vals.Length;

		var output = new object[ Math.Max( arrayLength, isArrayFunction ? 2 : 1 ), 1 ];
		for ( var i = 0; i < arrayLength; i++ )
		{
			output[ i, 0 ] = i < vals.Length ? vals[ i ] : ExcelError.ExcelErrorNA;
		}

		if ( isArrayFunction && vals.Length == 1 )
		{
			output[ 1, 0 ] = ExcelError.ExcelErrorNA;
		}

		return output;
	}

	[ExcelFunction( Category = "Formatting", Description = "Returns an array of values by splitting the provided delimitted string.  If index is provided, returns a single value at the specified 1 based array index.", IsMacroType = true )]
	public static object BTRSplit(
		[ExcelArgument( "Delimmited list of values to split into an array." )]
		string listValues,
		[ExcelArgument( "Optional.  Single character used to delimit the list.  Comma is the default." )]
		object? delimiter = null,
		[ExcelArgument( "Optional.  If BTRSplit is not sheet array formula, but used inline in cell formula, can provide 1 based index of value to return to eliminate need of INDEX() call." )]
		object? index = null,
		[ExcelArgument( "Optional.  If index is provided and is out of range of the values array, return this as the value.  Default is to return #VALUE." )]
		object? indexOutOfRangeDefault = null )
	{
		var delimiterArg = delimiter.Check( nameof( delimiter ), "," )[ 0 ];
		var indexArg = index.Check<int?>( nameof( index ), null );

		var vals = listValues.FromDelimitedString( delimiterArg );

		if ( vals.Length == 0 || ( indexArg != null && ( indexArg.Value < 1 || indexArg.Value > vals.Length ) ) )
		{
			return new object[,] { { ExcelError.ExcelErrorValue } };
		}
		else
		{
			var isArrayFunction = DnaApplication.IsArrayFormula;
			var output = new object[ Math.Max( vals.Length, isArrayFunction ? 2 : 1 ), 1 ];

			for ( var i = 0; i < vals.Length; i++ )
			{
				output[ i, 0 ] = vals[ i ];
			}

			if ( isArrayFunction && vals.Length == 1 )
			{
				output[ 1, 0 ] = ExcelError.ExcelErrorNA;
			}

			if ( !isArrayFunction && indexArg != null )
			{
				if ( indexArg.Value > output.GetLength( 0 ) )
				{
					var indexOutOfRangeArg = indexOutOfRangeDefault.Check<string?>( nameof( indexOutOfRangeDefault ), null );
					if ( indexOutOfRangeArg != null )
					{
						return indexOutOfRangeArg;
					}
				}

				return output[ indexArg.Value - 1, 0 ].ToString()!;
			}
			else
			{
				return output;
			}
		}
	}

	[ExcelFunction( Category = "Formatting", Description = "Returns localized resource string." )]
	public static string BTRResourceString(
		[ExcelArgument( "The resource key indicating which string to return." )]
		string key,
		[ExcelArgument( "Range of cells to search.  The first row must be column headers with first column of key and additional columns for each culture or culture-subculture containing the values." )]
		object[,] resourceStrings,
		[ExcelArgument( "Optional. The culture name to lookup.  If not provided, en-US is the default." )]
		object? cultureName = null,
		[ExcelArgument( "Optional. The name of the key column.  If not provided, 'key' is the default." )]
		object? keyName = null
	)
	{
		var cultureNameArg = cultureName.Check( nameof( cultureName ), "en-US" );

		var culturesToCheck = new[]
		{
			cultureNameArg,
			cultureNameArg.IndexOf( "-" ) > -1 ? cultureNameArg[ ..cultureNameArg.IndexOf( "-" ) ] : null,
			string.Compare( cultureNameArg, "en-US", true ) != 0 ? "en-US" : null,
			string.Compare( cultureNameArg.Split( '-' )[ 0 ], "en", true ) != 0 ? "en" : null
		}.Where( c => c != null );

		foreach ( var c in culturesToCheck )
		{
			var returnValues = Utility.LookupValues(
				resourceStrings,
				keyName.Check( nameof( keyName ), "key" ),
				c,
				new[] { key },
				null,
				true,
				ExcelEmpty.Value
			);

			if ( returnValues[ 0, 0 ] != null )
			{
				return (string)returnValues[ 0, 0 ]!;
			}
		}

		return key;
	}
}