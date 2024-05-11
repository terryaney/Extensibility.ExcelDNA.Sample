using System.Xml.Linq;
using System.Xml.XPath;
using KAT.Camelot.RBLe.Core.Calculations.Functions;

#pragma warning disable CA1822 // Mark members as static

namespace KAT.Camelot.Extensibility.Excel.AddIn;

class XmlContext
{
	public required XElement Element { get; init; }
	
	public DateTime MapToDate( string value, object? _ = null ) =>
		DateTime.FromOADate( Validation.ParseDate( value, "en-US" ) );
	
	public double MapToNumber( string value, object? _ = null ) => 
		Validation.ParseDecimal( value, "en-US" );

	public string? MapValue( string xpath, object? _ = null )
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

	public int MapOrdinal( int scopeDepth = 1, object? _ = null )
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

	public string MapFormatValue( object? value, string format, object? _ )
	{
		if ( value == null ) return string.Empty;

		var parsed = value is string s ? ParseValue( s ) : value;

		return Type.GetTypeCode( parsed.GetType() ) switch
		{
			TypeCode.String => (string)parsed,
			TypeCode.Int16 or TypeCode.Int32 => ( (int)parsed ).ToString( format ),
			TypeCode.Int64 => ( (long)parsed ).ToString( format ),
			TypeCode.UInt16 or TypeCode.UInt32 => ( (uint)parsed ).ToString( format ),
			TypeCode.UInt64 => ( (ulong)parsed ).ToString( format ),
			TypeCode.Single => ( (float)parsed ).ToString( format ),
			TypeCode.Double => ( (double)parsed ).ToString( format ),
			TypeCode.DateTime => ( (DateTime)parsed < new DateTime( 1900, 3, 1 ) ? ( (DateTime)parsed ).AddDays( 1 ) : (DateTime)parsed ).ToString( format ),
			_ => string.Empty,
		};
	}

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
