using System.Text.RegularExpressions;

namespace KAT.Camelot.Extensibility.Excel.AddIn.DataExport;

class ExportField
{
	// Format of input
	//		Help from: https://blog.mariusschulz.com/2014/06/03/why-using-in-regular-expressions-is-almost-never-what-you-actually-want
	//
	//	Value				Label					Description
	//	field				field					Just a regular field header
	//	field				field/delete			Whether or not to delete field, 1 to delete
	//  field				field/export			Whether or not to export field, 0 to skip
	//	[Status:2015]status	[Status:2015]status		Regular column based history field header
	//	[Status:2015]		[Status:2015]/export	Whether or not to ignore all columns with this prefix, 0 to skip
	//	[Status:2015]		[Status:2015]/delete	Whether or not to delete 2015 row, 1 to delete
	//	[Status]			[Status]/delete			Whether or not to delete all status, 1 to delete
	//	[Profile]			[Profile]/delete		Whether or not to delete entire profile, 1 to delete
	//	[Profile]			[Profile]/export		Whether or not to export entire profile, 0 to skip
	//	null				/delete					Whether or not to delete current row based history row, 1 to delete
	//	null				/export					Whether or not to export current row based history row, 0 to skip

	private static readonly Regex parse = new( @"(\[(?<table>[^:\]]+)(:(?<index>[^\]]+))?\])?(?<field>\S+)?", RegexOptions.Compiled );
	/*
		(\[(?<table>[^:\]]+)(:(?<index>[^\]]+))?\])?	- optional attempt at [table] and [table:index]
			\[											- opening [
				(?<table>[^:\]]+)						- match everything up until either : or ]
				(:(?<index>[^\]]+))?					- optional attempt at :index
					:									- match :
					(?<index>[^\]]+)					- match everything up until ]
			\]											- closing [
		(?<field>\S+)?									- optional attempt at field
			\S+											- match everything up until whitespace/eof
	*/

	public string? Field { get; set; }
	public string? HistoryTable { get; set; }
	public string? Index { get; set; }
	public string? NewIndex { get; set; }
	public bool DynamicIndex { get; set; }

	public static ExportField Parse( string info, Func<string, string, string>? indexReplace )
	{
		if ( string.IsNullOrEmpty( info ) ) return new ExportField();

		var parseInfo = parse.Match( info.Replace( ":New", ":Unique" ).Replace( " ", "_" ) );

		static string? groupValue( Group g ) => g.Success ? g.Value : null;

		var index = groupValue( parseInfo.Groups[ "index" ] );
		var table = groupValue( parseInfo.Groups[ "table" ] );
		// Tahiti supports these 'special' index tokens.  4.5 Allows New* or Unique*
		var tahitiIndexToken = index != null && ( index.StartsWith( "{UseDefault" ) || index.StartsWith( "{UseCalculated" ) );
		var dynamicIndex = tahitiIndexToken;

		if ( index != null && index.StartsWith( "{" ) && index.EndsWith( "}" ) && !tahitiIndexToken )
		{
			index = indexReplace!( table!, index );
			dynamicIndex = true;
		}

		return new ExportField
		{
			Field = groupValue( parseInfo.Groups[ "field" ] ),
			HistoryTable = table,
			Index = index,
			DynamicIndex = dynamicIndex
		};
	}

	public override string ToString() => ToString( FieldFormatType.FullSpecification );
	public string ToString( FieldFormatType formatType ) => string.Format( "{0}{1}", TableInfoKey, formatType == FieldFormatType.FullSpecification ? Field : null );

	public string? TableInfoKey => HistoryTable != null ? string.Format( "[{0}{1}]", HistoryTable, Index != null ? ":" + Index : null ) : null;
}
