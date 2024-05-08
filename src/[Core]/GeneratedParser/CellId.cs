using System.Text.RegularExpressions;

namespace KAT.Camelot.Extensibility.Excel.AddIn.GeneratedParser;

public struct CellId
{
	private static Regex columnId = new Regex( @"([a-zA-Z]+)(?=[0-9]+)" );
	private static Regex rowId = new Regex( @"(?<=[a-zA-Z]*)([0-9]+)" );

	private int column;
	private int row;

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
