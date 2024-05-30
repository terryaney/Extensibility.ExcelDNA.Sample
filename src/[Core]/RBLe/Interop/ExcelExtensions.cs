using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn.RBLe.Interop;

internal static class ExcelExtensions
{
	/// <summary>
	/// Given a 2 dimensional object array, it converts it into T array for typed used.
	/// </summary>
	/// <typeparam name="T"></typeparam>
	/// <param name="data"><see cref="object"/> parameter passed in that is either of type <see cref="double"/> or <see cref="object[,]"/>.</param>
	/// <returns>An T[,] array of the same dimensions as <paramref name="data"/>'s object[,] array.</returns>
	/// <remarks>
	/// <para>If <paramref name="data"/> is double, then a new double[,] array is constructed with only that element in it.</para>
	/// <para>This is also present in Evolution.CalcEngine.RBLeMacros.cs</para>
	/// </remarks>
	public static T[,] GetArray<T>( this MSExcel.Range reference )
	{
		var type = typeof( T );

		// Only ask for string, DateTime, int, double, if double/DateTime, those are 'native'
		// values range.Value returns those so can safely cast, if string or int, can't cast
		// possible double or DateTime to those, so use ChangeType
		var allowCast = typeof( string ) != type && typeof( int ) != type;

		var data = reference.Value;
		var array = new InteropArray( data as object[,] ?? new object[,] { { data } } );

		var rows = array.RowCount;
		var cols = array.ColumnCount;

		var result = new T[ rows, cols ];

		// Process the values
		for ( var i = 0; i < rows; i++ )
		{
			for ( var j = 0; j < cols; j++ )
			{
				var v = array[ i, j ];

				try
				{
					result[ i, j ] = (
						( v == null ) ? default :
						( allowCast ) ? (T)v :
						(T)Convert.ChangeType( v, type )
					)!;
				}
				catch ( InvalidOperationException ex )
				{
					throw new ApplicationException( string.Format( "{0}: {1}", ( reference.Cells[ i, 0 ] as MSExcel.Range )!.Address, ex.Message ), ex );
				}
			}
		}
		return result;
	}

	/// <summary>
	/// Given a 2 dimensional object array, it converts it into single dimensional T array for typed used.
	/// </summary>
	public static T[] GetValues<T>( this MSExcel.Range reference )
	{
		var type = typeof( T );
		var allowCast = typeof( string ) != type && typeof( int ) != type;
		var data = reference.Value;
		var array = new InteropArray( data as object[,] ?? new object[,] { { data } } );

		var rows = array.RowCount;
		var columns = array.ColumnCount;
		var size = Math.Max( rows, columns );
		var isVertical = rows == size;

		var result = new T[ size ];


		for ( var i = 0; i < size; i++ )
		{
			var v = isVertical
				? array[ i, 0 ]
				: array[ 0, i ];

			try
			{
				result[ i ] = (
					( v == null ) ? default :
					allowCast ? (T)v :
					(T)Convert.ChangeType( v, type )
				)!;
			}
			catch ( InvalidOperationException ex )
			{
				throw new ApplicationException( string.Format( "{0}: {1}", ( reference.Cells[ i, 0 ] as MSExcel.Range )!.Address, ex.Message ), ex );
			}
		}
		return result;
	}

	public static void SetArray<T>( this MSExcel.Range target, T[] array )
	{
		var vertical = new object[ array.Length, 1 ];
		for ( var i = 0; i < array.Length; i++ )
		{
			vertical[ i, 0 ] = array[ i ]!;
		}
		target.Value = vertical;
	}
}