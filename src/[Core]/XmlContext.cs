using System.Xml.Linq;
using System.Xml.XPath;
using KAT.Camelot.RBLe.Core.Calculations.Functions;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

class XmlContext
{
	public required XElement Element { get; init; }

	public DateTime BTRParseDate( string value ) => BTRParseDate( value, "en-US", null );
	public DateTime BTRParseDate( string value, string culture ) => BTRParseDate( value, culture, null );

	public DateTime BTRParseDate( string value, string culture, string? allowedFormats )
	{
		var formats = string.IsNullOrEmpty( allowedFormats )
			? null
			: allowedFormats.Split( '|' );

		return DateTime.FromOADate( Validation.ParseDate( value, culture ?? "en-US", formats ) );
	}

	public double BTRParseInteger( string value ) => BTRParseInteger( value, "en-US" );
	public double BTRParseInteger( string value, string culture ) => Validation.ParseInteger( value, culture ?? "en-US" );

	public double BTRParseDecimal( string value ) => BTRParseDecimal( value, "en-US" );
	public double BTRParseDecimal( string value, string culture ) => Validation.ParseDecimal( value, culture );

	public string? BTRGetMappingValue( string xpath )
	{
		try
		{
			if ( Element.XPathEvaluate( xpath ) is IEnumerable<object> enumerable )
			{
				var node = enumerable.First();

				if ( node != null )
				{
					return ( node is XElement element ) ? element.Value : ( (XAttribute)node ).Value;
				}
			}

			return null;
		}
		catch ( Exception ex )
		{
			throw new ApplicationException( $"GetMappingValue failed.  Current xpath = {xpath}", ex );
		}
	}

	public int BTRGetMappingOrdinal() => BTRGetScopedMappingOrdinal( 1 );
	public int BTRGetScopedMappingOrdinal( int scopeDepth )
	{
		if ( scopeDepth <= 1 )
		{
			return Element.ElementsBeforeSelf( Element.Name ).Count() + 1;
		}
		else
		{
			var scope = Element;
			var scopeXPath = scope.Name.LocalName;

			// Loop one less than depth so don't include 'scope' elements name in xpath
			for ( var i = 1; i < scopeDepth; i++ )
			{
				scope = scope.Parent;
				scopeXPath = scope!.Name.LocalName + "/" + scopeXPath;
			}

			// Now back up once more to get to proper scope element before xpath execution
			scope = scope.Parent;

			return scope!.XPathSelectElements( scopeXPath ).TakeWhile( d => d != Element ).Count() + 1;
		}
	}

	public string BTRNumberFormat( double value, string format ) => BTRNumberFormat( value, format, "en-US" );
	public string BTRNumberFormat( double value, string format, string culture ) =>  Utility.LocaleFormat( value, format, culture );

	public string BTRDateFormat( DateTime value, string format ) => BTRDateFormat( value, format, "en-US" );
	public string BTRDateFormat( DateTime value, string format, string culture ) => Utility.LocaleFormat( ( value < new DateTime( 1900, 3, 1 ) ? value.AddDays( 1 ) : value ), format, culture );
	

	public object ChangeType( object value, Type type ) => Convert.ChangeType( value, type );

	private static object ParseValue( string value )
	{
		if ( int.TryParse( value, out var iValue ) ) return iValue;

		if ( long.TryParse( value, out var lValue ) ) return lValue;

		if ( double.TryParse( value, out var dValue ) ) return dValue;

		if ( DateTime.TryParse( value, out var dtValue ) ) return dtValue;

		return value;
	}
}
