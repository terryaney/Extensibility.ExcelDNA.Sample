namespace KAT.Camelot.Extensibility.Excel.AddIn;

/// <summary>
/// Wrapper for Microsoft Interop Arrays that are 1 based to make accessing them
/// via C# 0 based arrays work.
/// </summary>
public class InteropArray
{
	private readonly object[,] data;

	public int RowCount { get; private set; }
	public int ColumnCount { get; private set; }

	private readonly int rowOffset;
	private readonly int colOffset;

	public InteropArray( object[,] data )
	{
		this.data = data;
		RowCount = data.GetLength( 0 );
		ColumnCount = data.GetLength( 1 );

		// If GetArray is called on single cell, c# creates the
		// array instead of Excel and it is 0 based.
		rowOffset = data.GetLowerBound( 0 );
		colOffset = data.GetLowerBound( 1 );
	}

	public IEnumerable<object[]> Rows
	{
		get
		{
			for ( var i = 0; i < RowCount; i++ )
			{
				var row = new object[ ColumnCount ];
				for ( var j = 0; j < ColumnCount; j++ )
				{
					row[ j ] = data[ i + rowOffset, j + colOffset ];
				}
				yield return row;
			}
		}
	}

	public object this[ int x, int y ] => data[ x + rowOffset, y + colOffset ];
}