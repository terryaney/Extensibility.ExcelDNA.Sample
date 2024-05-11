using System.Text.RegularExpressions;

namespace KAT.Camelot.Extensibility.Excel.AddIn.GeneratedParser;

public readonly struct CellId
{
	private static readonly Regex columnId = new( @"([a-zA-Z]+)(?=[0-9]+)" );
	private static readonly Regex rowId = new( @"(?<=[a-zA-Z]*)([0-9]+)" );

	private readonly int column;
	private readonly int row;

	public int Column => column;
	public int Row => row;

	public CellId( string id )
		: this(
			columnId.Match( id ).Groups[ 0 ].Value.ToUpper(),
			int.Parse( rowId.Match( id ).Groups[ 0 ].Value ) )
	{
	}

	public CellId( string columnId, int row )
		: this( columnId.ToUpper().FromBase26(), row )
	{
	}

	public CellId( int column, int row )
	{
		this.column = column;
		this.row = row;
	}

	public string Name { get { return ToString(); } }

	public override string ToString()
	{
		return GetName( Column, Row );
	}

	public static string GetName( int column, int row )
	{
		return column.ToBase26() + row.ToString();
	}
}
