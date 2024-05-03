using ExcelDna.Integration;
using KAT.Camelot.RBLe.Core;
using KAT.Camelot.RBLe.Core.Calculations;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

partial class ExcelApi
{
	public static string SheetName( this ExcelReference reference ) => (string)XlCall.Excel( XlCall.xlfGetCell, (int)GetCellType.SheetRef, reference );

	public static string? GetText( this ExcelReference cell )
	{
		var value = cell.GetValue();
		return value.Equals( ExcelEmpty.Value ) ? null : (string)XlCall.Excel( XlCall.xlfGetCell, (int)GetCellType.Text, cell );
	}

	public static string? GetFormula( this ExcelReference cell )
	{
		var f = XlCall.Excel( XlCall.xlfGetCell, (int)GetCellType.Formula, cell );
		var formula = f is ExcelError check && check == ExcelError.ExcelErrorValue ? null : (string)f;
		return !string.IsNullOrEmpty( formula ) ? formula : null;
	}

	public static string GetAddress( this ExcelReference? reference )
	{
		try
		{
			var address = (string)XlCall.Excel( XlCall.xlfReftext, reference, true /* true - A1, false - R1C1 */ );
			return address;
		}
		catch ( Exception ex )
		{
			throw new ApplicationException( $"GetAddress failed.  reference.RowFirst:{reference?.RowFirst}, reference.RowLast:{reference?.RowLast}, reference.ColumnFirst:{reference?.ColumnFirst}, reference.ColumnLast:{reference?.ColumnLast}", ex );
		}
	}

	public static ExcelReference Extend( this ExcelReference start, ExcelReference end )
	{
		return new ExcelReference(
			Math.Min( start.RowFirst, end.RowFirst ),
			Math.Max( start.RowLast, end.RowLast ),
			Math.Min( start.ColumnFirst, end.ColumnFirst ),
			Math.Max( start.ColumnLast, end.ColumnLast ),
			start.SheetId );
	}

	public static ExcelReference End( this ExcelReference reference, DirectionType direction, bool ignoreEmpty = false )
	{
		ExcelReference end = null!;

		reference.RestoreSelection( () =>
		{
			var value = 
				ignoreEmpty ? ExcelEmpty.Value :
				direction == DirectionType.Down ? reference.Offset( 1, 0 ).GetValue() :
				direction == DirectionType.ToRight ? reference.Offset( 0, 1 ).GetValue() :
				direction == DirectionType.ToLeft ? reference.Offset( 0, -1 ).GetValue() :
				/* DirectionType.Up */                reference.Offset( -1, 0 ).GetValue();

			var isEmpty = value == ExcelEmpty.Value || ( value.GetType() == typeof( string ) && string.IsNullOrEmpty( (string)value ) );

			if ( !ignoreEmpty && isEmpty )
			{
				end = reference;
			}
			else
			{
				// Govert talks about 'messiness' at http://stackoverflow.com/a/10920622/166231, but pretty straight forward
				XlCall.Excel( XlCall.xlcSelectEnd, (int)direction );

				var selection = ( XlCall.Excel( XlCall.xlfSelection ) as ExcelReference )!;
				var row = selection.RowFirst;
				var col = selection.ColumnFirst;

				end = new ExcelReference( row, row, col, col, selection.SheetId );
			}
		} );

		return end;
	}

	public static ExcelReference Offset( this ExcelReference reference, int rows, int cols )
	{
		return new ExcelReference(
			reference.RowFirst + rows,
			reference.RowLast + rows,
			reference.ColumnFirst + cols,
			reference.ColumnLast + cols,
			reference.SheetId );
	}

	public static ExcelReference Corner( this ExcelReference reference, CornerType corner )
	{
		var row = corner == CornerType.UpperLeft || corner == CornerType.UpperRight
			? reference.RowFirst
			: reference.RowLast;

		var column = corner == CornerType.UpperLeft || corner == CornerType.LowerLeft
			? reference.ColumnFirst
			: reference.ColumnLast;

		return new ExcelReference( row, row, column, column, reference.SheetId );
	}

	public static T?[,] GetArray<T>( this ExcelReference reference )
	{
		// Information about decisions made on how to read bulk data
		//		http://stackoverflow.com/questions/17359835/what-is-the-difference-between-text-value-and-value2
		//		https://fastexcel.wordpress.com/2011/11/30/text-vs-value-vs-value2-slow-text-and-how-to-avoid-it/

		var type = typeof( T );

		// Only ask for string, DateTime, int, double, if double/DateTime, those are 'native'
		// return range.Value return those so can safely cast, if string or int, can't cast
		// possible double or DateTime to those, so use ChangeType
		var allowCast = typeof( string ) != type && typeof( int ) != type;

		var data = reference.GetValueArray();

		var rows = data.RowCount;
		var cols = data.ColumnCount;

		var result = new T?[ rows, cols ];

		// Process the values
		for ( var i = 0; i < rows; i++ )
		{
			for ( var j = 0; j < cols; j++ )
			{
				var v = data[ i, j ];

				try
				{
					result[ i, j ] = FromInteropValue<T>( type, v, allowCast );
				}
				catch ( InvalidOperationException ex )
				{
					throw new ApplicationException( string.Format( "{0}: {1}", reference.Corner( CornerType.UpperLeft ).Offset( i, j ).GetAddress(), ex.Message ), ex );
				}
			}
		}
		return result;
	}

	public static InteropArray GetValueArray( this ExcelReference reference )
	{
		// Following call returns object[] array containing string, DateTime (correctly preserving Dates), and double values.
		// As reply to me states, I do not lose precision on currency formatted cells going to .NET
		// https://fastexcel.wordpress.com/2011/11/30/text-vs-value-vs-value2-slow-text-and-how-to-avoid-it/#comment-3521

		// Unforunately, I can't use ExcelDNA reference as they have no method to return a value as DateTime...

		// var data = reference.GetValue() as object[,]; // dates converted to doubles, no way to get back
		var data = reference.GetRange().Value;
		// Reference was only 1 cell if null, so not an object[,] but actual value but change to array
		var value = data as object[,] ?? new object[,] { { data } };
		return new InteropArray( value );
	}

	private static T? FromInteropValue<T>( Type type, object value, bool allowCast )
	{
		if ( value == null ) return default;

		value.ThrowOnInteropError( value.GetType() );

		if ( allowCast ) return (T)value;

		return (T)Convert.ChangeType( value, type );
	}

	public static void ThrowOnInteropError( this object value, Type type )
	{
		if ( typeof( int ) == type )
		{
			var i = (int)value;

			if ( i == -2146826288d ) throw new ExcelErrorException( "ExcelError.ExcelErrorNull", "#NULL!" );
			if ( i == -2146826281d ) throw new ExcelErrorException( "ExcelError.ExcelErrorDiv0", "#DIV/0!" );
			if ( i == -2146826265d ) throw new ExcelErrorException( "ExcelError.ExcelErrorRef", "#REF!" );
			if ( i == -2146826259d ) throw new ExcelErrorException( "ExcelError.ExcelErrorName", "#NAME?" );
			if ( i == -2146826252d ) throw new ExcelErrorException( "ExcelError.ExcelErrorNum", "#NUM!" );
			if ( i == -2146826246d ) throw new ExcelErrorException( "ExcelError.ExcelErrorNA", "#N/A" );
			if ( i == -2146826273d ) throw new ExcelErrorException( "ExcelError.ExcelErrorValue", "#VALUE!" );
		}
	}
}