using ExcelDna.Integration;

namespace KAT.Camelot.Extensibility.Excel.AddIn.Functions;

public static class XmlMapping
{
	[ExcelFunction( Category = "Xml Mapping", Description = "Returns the current ordinal position of the current mapping element being processed.  If 'scopeDepth' is passed, it is the current ordinal position of the ancestor mapping element determined by 'scopeDepth' levels.  Placeholder returning defaultValue in Excel." )]
	public static int MapOrdinal( 
		[ExcelArgument( "How many parent levels to walk back up to determine mapping scope.  Default value is one." )] object? scopeDepth = null,
		[ExcelArgument( "Value to return to make coding specification formulas easier." )] object? defaultValue = null
	) => defaultValue.Check( nameof( defaultValue ), 1 );

	[ExcelFunction( Category = "Xml Mapping", Description = "Returns 'fieldName' (element or attribute) value from provided Xml datasource.  Placeholder returning defaultValue in Excel." )]
	public static string MapValue( 
		[ExcelArgument( "The name of the xml element or attribute." )] string fieldName, 
		[ExcelArgument( "Value to return to make coding specification formulas easier." )] object? defaultValue = null 
	) => defaultValue.Check( nameof( defaultValue ), fieldName );
	
	[ExcelFunction( Category = "Xml Mapping", Description = "Converts a value into a number.  Needed for 'strong typing' in calculated expressions." )]
	public static double MapToNumber(
		[ExcelArgument( "The value to attempt to convert to a number." )] object value,
		[ExcelArgument( "Value to return to make coding specification formulas easier." )] object? defaultValue = null )
	{
		var defaultValueArg = defaultValue.Check<double?>( nameof( defaultValue ), null );

		if ( ( value is string v && double.TryParse( v, out var d ) ) ) return d;

		try
		{
			return Convert.ToDouble( value );
		}
		catch
		{
			return defaultValueArg ?? 0d;
		}
	}

	[ExcelFunction( Category = "Xml Mapping", Description = "Converts a value into a date.  Needed for 'strong typing' in calculated expressions." )]
	public static DateTime MapToDate(
		[ExcelArgument( "The value to attempt to convert to a date." )] object value,
		[ExcelArgument( "Value to return to make coding specification formulas easier." )] object? defaultValue = null 
	)
	{
		var defaultValueArg = defaultValue.Check<DateTime?>( nameof( defaultValue ), null );

		if ( ( value is string v && DateTime.TryParse( v, out var d ) ) ) return d;

		try
		{
			return Convert.ToDateTime( value );
		}
		catch
		{
			return defaultValueArg ?? DateTime.Now;
		}
	}

	[ExcelFunction( Category = "Xml Mapping", Description = "Returns 'value' formatted as string given the desired 'format' pattern.  Placeholder returning [FormatValue('value', 'format'] string." )]
	public static string MapFormatValue(
		[ExcelArgument( "The value (or model field) to format." )] object? value,
		[ExcelArgument( "A valid C# format string in the format of {0:format}." )] string format,
		[ExcelArgument( "Value to return to make coding specification formulas easier." )] object? defaultValue = null 
	)
	{
		var defaultValueArg = defaultValue.Check( nameof( defaultValue ), $"[MapFormatValue({value}, \"{format.Replace( "\"", "\\\"" )}\")]" );
		if ( value == null ) return defaultValueArg;

		var parsed = value is string s ? ParseValue( s ) : value;

		return Type.GetTypeCode( parsed.GetType() ) switch
		{
			TypeCode.String => (string)parsed,
			TypeCode.Int16 or TypeCode.Int32 => ( (int)parsed ).ToString( format ),
			TypeCode.Int64 => ( (long)parsed ).ToString( format ),
			TypeCode.UInt16 or TypeCode.UInt32 => ( (uint)parsed ).ToString( format ),
			TypeCode.UInt64 => ( (ulong)parsed ).ToString( format ),
			TypeCode.Single => ( (float)parsed ).ToString( format ),
			TypeCode.Double => new[] { "m", "d", "y", "h", "s" }.Any( s => format.Contains( s, StringComparison.OrdinalIgnoreCase ) )
				? ( DateTime.FromOADate( (double)parsed ) ).ToString( format )
				: ( (double)parsed ).ToString( format ),
			TypeCode.DateTime => ( (DateTime)parsed ).ToString( format ),
			_ => defaultValueArg,
		};
	}

	private static object ParseValue( string value )
	{
		if ( int.TryParse( value, out var iValue ) ) return iValue;
		if ( long.TryParse( value, out var lValue ) ) return lValue;
		if ( double.TryParse( value, out var dValue ) ) return dValue;
		if ( DateTime.TryParse( value, out var dtValue ) ) return dtValue;

		return value;
	}
}