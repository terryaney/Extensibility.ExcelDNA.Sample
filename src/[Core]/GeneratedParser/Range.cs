namespace KAT.Camelot.Extensibility.Excel.AddIn.GeneratedParser;

public struct Range
{
	public CellId FromId { get; set; }
	private CellId ToId { get; set; }

	public Range( CellId fromId, CellId toId ) : this()
	{
		this.FromId = fromId;
		this.ToId = toId;
	}

	public Range( IEnumerable<CellId> cellIds ) : this()
	{
		var ids = cellIds.Take( 2 ).ToArray();

		this.FromId = ids[ 0 ];
		this.ToId = ( ids.Length > 1 ) ? ids[ 1 ] : ids[ 0 ];
	}

	public readonly IEnumerable<CellId> Cells => GetCells( FromId, ToId );

	public static IEnumerable<CellId> GetCells( CellId fromId, CellId toId )
	{
		return from c in Enumerable.Range( fromId.Column, toId.Column - fromId.Column + 1 )
			   from r in Enumerable.Range( fromId.Row, toId.Row - fromId.Row + 1 )
			   select new CellId( c, r );
	}
}