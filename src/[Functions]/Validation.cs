using ExcelDna.Integration;
using KAT.Camelot.RBLe.Core.Calculations;
using KAT.Camelot.RBLe.Core.Calculations.Functions;

namespace KAT.Camelot.Extensibility.Excel.AddIn.Functions;

public static class DnaValidation
{
	[KatExcelFunction(
		Category = "Validation",
		Description = "Returns whether the provided input is a valid US banking routing number.",
		Remarks = "The logic for this method was created by following [this routine](http://www.wikihow.com/Calculate-the-Check-Digit-of-a-Routing-Number-from-an-Illegible-Check)."
	)]
	public static bool BTRValidateRoutingNumber( 
		[ExcelArgument( "The routing number input provided by user to validate." )] 
		string value 
	) => Validation.ValidateRoutingNumber( value );
	

	[ExcelFunction( Category = "Validation", Description = "Returns whether the regular expression finds a match in the input string." )]
	public static bool BTRValidateRegEx(
		[ExcelArgument( "The input provided by user to search for a match." )] 
		string value,
		[ExcelArgument( "The regular expression pattern to match." )] 
		string pattern 
	) => Validation.ValidateRegEx( value, pattern );
	

	[KatExcelFunction( 
		Category = "Validation", 
		Description = "Validates and converts the input string representation of a date and time, supporting culture specific formats, to its Date equivalent.  Throws an exception if not a valid date.",
		Exceptions = "`ArgumentOutOfRangeException` - `value` does not contain a valid string representation of a date and time."
	)]
	public static object BTRParseDate(
		[ExcelArgument( "A string that contains a date and time to convert." )]
		string value,
		[KatExcelArgument(
			Description = "A string that supplies culture-specific format information about `value`.",
			Type = typeof( string ),
			Default = "en-US"
		)]
		object? culture = null,
		[ExcelArgument( "A `|` delimitted string that supplies a list of allowable formats to attempt to parse `value`." )]
		string? allowedFormats = null,
		[ExcelArgument( "A `,` delimitted string of allowable dates to validate in the format of 1..N, Last, Mon-Sun, Mon-Sun.[N|Last] (Nth occurence of or last day in month), or W1-W5 (first through the fifth week of month).  If the 'pattern' starts with '!' it is a 'not' check." )]
		string? validDates = null )
	{
		var cultureArg = culture.Check( nameof( culture ), "en-US" );

		var result = Validation.TryParseDate(
			value,
			cultureArg,
			string.IsNullOrEmpty( allowedFormats ) ? null : allowedFormats.Split( '|' ),
			string.IsNullOrEmpty( validDates ) ? null : validDates.Split( ',' )
		);

		return (object?)result ?? ExcelError.ExcelErrorValue;
	}

	[ExcelFunction( Category = "Validation", Description = "Returns an integer indicating whether an input string is a valid date and within specified range. -1 if invalid, 0 if valid and in range, and 1 if outside of range." )]
	public static double BTRValidateDate(
		[ExcelArgument( "A string that contains a date and time to convert." )]
		string value,
		[ExcelArgument( "A DateTime representing the minimum value allowed if `value` is a date." )]
		DateTime minimum,
		[ExcelArgument( "A DateTime representing the minimum value allowed if `value` is a date." )]
		DateTime maximum,
		[KatExcelArgument(
			Description = "A string that supplies culture-specific format information about `value`.",
			Type = typeof( string ),
			Default = "en-US"
		)]
		object? culture = null,
		[ExcelArgument( "A `|` delimitted string that supplies a list of allowable formats to attempt to parse `value`." )]
		string? allowedFormats = null,
		[ExcelArgument( "A `,` delimitted string of allowable dates to validate in the format of 1..N, Last, Mon-Sun, Mon-Sun.[N|Last] (Nth occurence of or last day in month), or W1-W5 (first through the fifth week of month).  If the 'pattern' starts with '!' it is a 'not' check." )]
		string? validDates = null )
	{
		try
		{
			var cultureArg = culture.Check( nameof( culture ), "en-US" );

			var result = Validation.TryParseDate(
				value,
				cultureArg,
				string.IsNullOrEmpty( allowedFormats ) ? null : allowedFormats.Split( '|' ),
				string.IsNullOrEmpty( validDates ) ? null : validDates.Split( '|' ) );

			if ( result == null )
			{
				return -1;
			}

			var date = DateTime.FromOADate( result.Value );
			return date <= maximum && date >= minimum ? 0 : 1;
		}
		catch
		{
			return -1;
		}
	}

	[KatExcelFunction( 
		Category = "Validation", 
		Description = "Validates and converts the input string representation of an age/date, supporting culture specific formats, to its Date equivalent.  Throws an exception if not a valid age/date.",
		Exceptions = "`ArgumentOutOfRangeException` - `value` does not contain a valid string representation of a date and time."
	)]
	public static object BTRParseAgeDate(
		[ExcelArgument( "A string that contains a date or age to convert." )]
		string value,
		[ExcelArgument( "The participant's date of birth." )]
		DateTime dateBirth,
		[KatExcelArgument(
			Description = "Additional options to apply to date (FirstOfMonthOrCoincident=1, LastOfMonthOrCoincident=2).",
			Type = typeof( DateOptionsType ),
			Default = "DateOptionsType.None"
		)]
		object? dateOptions = null,
		[KatExcelArgument(
			Description = "A string that supplies culture-specific format information about `value`.",
			Type = typeof( string ),
			Default = "en-US"
		)]
		object? culture = null,
		[ExcelArgument( "A `|` delimitted string that supplies a list of allowable formats to attempt to parse `value`." )]
		string? allowedFormats = null )
	{
		//(Functions.CalculationContributionType)contributionType
		var cultureArg = culture.Check( nameof( culture ), "en-US" );
		var dateOptionsArg = (DateOptionsType)dateOptions.Check( nameof( dateOptions ), 0 );

		var formats = string.IsNullOrEmpty( allowedFormats )
			? null
			: allowedFormats.Split( '|' );

		var result = Validation.TryParseAgeDate( value, dateBirth, dateOptionsArg, cultureArg, formats );

		return (object?)result ?? ExcelError.ExcelErrorValue;
	}

	[ExcelFunction( Category = "Validation", Description = "Returns an integer indicating whether an input string is a valid age/date and within specified range. -1 if invalid, 0 if valid and in range, and 1 if outside of range." )]
	public static double BTRValidateAgeDate(
		[ExcelArgument( "A string that contains a date or age to validate." )]
		string value,
		[ExcelArgument( "The participant's date of birth." )]
		DateTime dateBirth,
		[ExcelArgument( "A DateTime representing the minimum value allowed if `value` is a date." )]
		DateTime minimum,
		[ExcelArgument( "A DateTime representing the minimum value allowed if `value` is a date." )]
		DateTime maximum,
		[KatExcelArgument(
			Description = "Additional options to apply to date (FirstOfMonthOrCoincident=1, LastOfMonthOrCoincident=2).",
			Type = typeof( DateOptionsType ),
			Default = "DateOptionsType.None"
		)]
		object? dateOptions = null,
		[KatExcelArgument(
			Description = "A string that supplies culture-specific format information about `value`.",
			Type = typeof( string ),
			Default = "en-US"
		)]
		object? culture = null,
		[ExcelArgument( "A `|` delimitted string that supplies a list of allowable formats to attempt to parse `value`." )]
		string? allowedFormats = null )
	{
		try
		{
			var cultureArg = culture.Check( nameof( culture ), "en-US" );
			var dateOptionsArg = (DateOptionsType)dateOptions.Check( nameof( dateOptions ), 0 );
			var formats = string.IsNullOrEmpty( allowedFormats )
				? null
				: allowedFormats.Split( '|' );

			var result = Validation.TryParseAgeDate( value, dateBirth, dateOptionsArg, cultureArg, formats );

			if ( result == null )
			{
				return -1;
			}

			var date = DateTime.FromOADate( result.Value );

			return date <= maximum && date >= minimum ? 0 : 1;
		}
		catch
		{
			return -1;
		}
	}

	[KatExcelFunction( 
		Category = "Validation", 
		Description = "Validates and converts the input string representation of a number to its integer (no decimals) equivalent.",
		Exceptions = "`ArgumentOutOfRangeException` - `value` does not contain a valid string representation of a an integer (whole number)."
	)]
	public static double BTRParseInteger(
		[ExcelArgument( "A string that contains a number to convert." )]
		string value,
		[KatExcelArgument(
			Description = "A string that supplies culture-specific format information about `value`.",
			Type = typeof( string ),
			Default = "en-US"
		)]
		object? culture = null )
	{
		var cultureArg = culture.Check( nameof( culture ), "en-US" );
		return Validation.ParseInteger( value, cultureArg );
	}

	[ExcelFunction( Category = "Validation", Description = "Returns an integer indicating whether an input string is a valid integer and within specified range. -1 if invalid, 0 if valid and in range, and 1 if outside of range." )]
	public static double BTRValidateInteger(
		[ExcelArgument( "A string that contains a date and time to convert." )]
		string value,
		[ExcelArgument( "A integer representing the minimum value allowed if `value` is a integer." )]
		int minimum,
		[ExcelArgument( "A integer representing the minimum value allowed if `value` is a integer." )]
		int maximum,
		[KatExcelArgument(
			Description = "A string that supplies culture-specific format information about `value`.",
			Type = typeof( string ),
			Default = "en-US"
		)]
		object? culture = null )
	{
		try
		{
			var cultureArg = culture.Check( nameof( culture ), "en-US" );
			var result = Validation.TryParseInteger( value, cultureArg );

			if ( result == null )
			{
				return -1;
			}

			return result.Value <= maximum && result.Value >= minimum ? 0 : 1;
		}
		catch
		{
			return -1;
		}
	}

	[KatExcelFunction( 
		Category = "Validation", 
		Description = "Validates and converts the input string representation of a number to its decimal equivalent.",
		Exceptions = "`ArgumentOutOfRangeException` - `value` does not contain a valid string representation of a an decimal number."
	)]
	public static double BTRParseDecimal(
		[ExcelArgument( "A string that contains a number to convert." )]
		string value,
		[KatExcelArgument(
			Description = "A string that supplies culture-specific format information about `value`.",
			Type = typeof( string ),
			Default = "en-US"
		)]
		object? culture = null )
	{
		var cultureArg = culture.Check( nameof( culture ), "en-US" );
		return Validation.ParseDecimal( value, cultureArg );
	}

	[ExcelFunction( Category = "Validation", Description = "Returns an integer indicating whether an input string is a valid decimal value and within specified range. -1 if invalid, 0 if valid and in range, and 1 if outside of range." )]
	public static double BTRValidateDecimal(
		[ExcelArgument( "A string that contains a date and time to convert." )]
		string value,
		[ExcelArgument( "A double representing the minimum value allowed if `value` is a double." )]
		double minimum,
		[ExcelArgument( "A double representing the minimum value allowed if `value` is a double." )]
		double maximum,
		[KatExcelArgument(
			Description = "A string that supplies culture-specific format information about `value`.",
			Type = typeof( string ),
			Default = "en-US"
		)]
		object? culture = null,
		[KatExcelArgument(
			Description = "An integer value representing the maximum number of decimal places allowed.",
			Type = typeof( int )
		)]
		object? decimalPlaces = null )
	{
		try
		{
			var cultureArg = culture.Check( nameof( culture ), "en-US" );
			var decimalPlacesArg = decimalPlaces.Check<int?>( nameof( decimalPlaces ), null );
			var result = Validation.TryParseDecimal( value, cultureArg, decimalPlacesArg );

			if ( result == null )
			{
				return -1;
			}

			return result.Value <= maximum && result.Value >= minimum ? 0 : 1;
		}
		catch
		{
			return -1;
		}
	}
}
