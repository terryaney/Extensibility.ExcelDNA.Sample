using System.Globalization;
using System.Text.Json.Nodes;
using ExcelDna.Integration;
using KAT.Camelot.Domain.Extensions;
using KAT.Camelot.Extensibility.Excel.AddIn.ExcelApi;
using KAT.Camelot.RBLe.Core.Calculations;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn.ConfigurationExport;

enum LookupConfigurationType
{
	DataTables,
	FrameworkTables,
	RateTables,
	GlobalTables
}

class GlobalTables
{
	public static JsonObject Export( MSExcel.Worksheet[] sheets )
	{
		var globalSpecifications = new JsonObject();

		if ( sheets.Length == 1 )
		{
			var propertyName = sheets[ 0 ].RangeOrNull<string>( Constants.SpecSheet.RangeNames.SheetType ) == Constants.SpecSheet.SheetTypes.GlobalLookupTables
				? "dataTables"
				: "rateTable";

			globalSpecifications[ propertyName ] = GetGlobalTables( sheets[ 0 ] );
		}
		else
		{
			foreach( var sheet in sheets )
			{
				var propertyName = sheet.Name == Constants.SpecSheet.TabNames.DataLookupTables
					? "dataTables"
					: "rateTables";
				globalSpecifications[ propertyName ] = GetGlobalTables( sheet );
			}
		}

		return globalSpecifications;
	}

	private static JsonObject GetGlobalTables( MSExcel.Worksheet worksheet )
	{
		var tables = GetGlobalLookupTables( worksheet ).ToJsonArray();
		var version = worksheet.RangeOrNull<string>( Constants.SpecSheet.RangeNames.SheetVersion );

		return new JsonObject
		{
			{ "version", version },
			{ "tables", tables }
		};
	}

	private static IEnumerable<JsonObject> GetGlobalLookupTables( MSExcel.Worksheet worksheet )
	{
		var configurationType = worksheet.RangeOrNull<string>( Constants.SpecSheet.RangeNames.SheetType ) switch
		{
			Constants.SpecSheet.SheetTypes.GlobalLookupTables => LookupConfigurationType.GlobalTables,
			Constants.SpecSheet.SheetTypes.GlobalRateTables or Constants.SpecSheet.SheetTypes.ClientRateTables => LookupConfigurationType.RateTables,
			_ => LookupConfigurationType.DataTables // TODO: Is this right?
		};

		var headerOffset = configurationType == LookupConfigurationType.DataTables ? 2 : 1;

		var firstColumn = 
			worksheet.Range[ worksheet.RangeOrNull<string>( Constants.SpecSheet.RangeNames.TableStartAddress ) ]
				.GetReference()
				.Offset( headerOffset, 0 );

		while ( !string.IsNullOrEmpty( firstColumn.GetValue<string>() ) )
		{
			var tableName = firstColumn.Offset( -headerOffset, 1 ).GetValue<string>()!;
			var lastColumn = firstColumn.End( DirectionType.ToRight );
			var tableInclude = firstColumn.Offset( -1, 1 ).GetValue<string>() ?? "Y";

			if ( configurationType != LookupConfigurationType.DataTables || tableInclude.StartsWith( "Y" ) )
			{
				var lastRow = firstColumn.End( DirectionType.Down );
				var data =
					firstColumn
						.Extend( lastColumn )
						.Extend( lastRow )
						.GetValueArray();

				if ( tableInclude.Contains( "/customize" ) )
				{
					throw new NotImplementedException( $"No support for /customize. {tableName} used this flag.  See if it is needed." );
				}

				yield return new JsonObject
				{
					{ "name", tableName },
					{ "columns", new JsonArray().AddItems( data.Rows.First() ) },
					{ "rows", data.Rows.Skip( 1 ).Select( r => new JsonArray().AddItems( r.Select( GetExportValue ), includeNulls: true ) ).ToJsonArray() }
				};
			}

			firstColumn = lastColumn.End( DirectionType.ToRight, ignoreEmpty: true );
		}
	}

	private static string? GetExportValue( object value )
	{
		if ( value == ExcelEmpty.Value ) return null;

		var d = value as double?;
		if ( d != null )
		{
			// .NET Core changed in the underlying implementation of the Double.ToString() method. 
			// In .NET Core, the method has been updated to produce a round-trippable result by default, 
			// which means it will always return a string that, when parsed, will produce the original number.
			// To reproduce what we had in .NET Framework it is suggested to use the G15 format.  
			// The "G" format specifier stands for "general", and it formats the number in the most compact, human-readable form.
			// In .NET Framework, the default precision for double.ToString() without any format specifier is up to 15 digits, 
			// which can be either to the left or right of the decimal point. This means it can include up to 15 significant digits, 
			// and the remaining digits are replaced with zeros.
			return d.Value.ToString( "G15", CultureInfo.InvariantCulture );

			// I was going to count the number of decimal places and use that + 6, but it seems to be a bit overkill.
			// var decimalValues = ( (int)Math.Floor( Math.Abs( d.Value ) ) ).ToString().Length;
			// return d.Value.ToString( $"G{decimalValues + 6}", CultureInfo.InvariantCulture );
		}

		var dt = value as DateTime?;
		if ( dt != null ) return dt.Value.ToString( "yyyy-MM-dd" );

		var s = (string)value;

		// Excel does all caps for these.
		if ( s == "TRUE" ) s = "true";
		else if ( s == "FALSE" ) s = "false";

		return s;
	}
}