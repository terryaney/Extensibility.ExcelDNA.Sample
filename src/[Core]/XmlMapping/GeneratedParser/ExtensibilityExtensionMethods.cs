using System.Text;

namespace KAT.Camelot.Extensibility.Excel.AddIn.XmlMappingExport.GeneratedParser;

public static class ExtensibilityExtensionMethods
{
	private static readonly string upperCase = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

	public static string ToBase26( this int value )
	{
		return ToBaseN( value, upperCase );
	}

	public static string ToBaseN( this int value, string characterSet )
	{
		var result = new StringBuilder();
		var n = characterSet.Length;

		do
		{
			var remainder = value % n;
			value /= n;
			result.Insert( 0, new[] { (char)( characterSet[ remainder ] ) } );
		}
		while ( value > 0 );

		return result.ToString();
	}

	public static int FromBase26( this string value )
	{
		return FromBaseN( value, upperCase );
	}

	public static int FromBaseN( this string value, string characterSet )
	{
		var n = characterSet.Length;

		return value.ToCharArray().Aggregate( 0, ( acc, c ) => ( acc * n ) + characterSet.IndexOf( c ) );
	}
}