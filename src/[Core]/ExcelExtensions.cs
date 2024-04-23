using ExcelDna.Integration.CustomUI;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Extensibility.Excel.AddIn;

public static class ExcelExtensions
{
	public static void InvalidateControls( this IRibbonUI ribbon, params string[] controlIds )
	{
		foreach ( var controlId in controlIds )
		{
			ribbon.InvalidateControl( controlId );
		}
	}

	public static MSExcel.Worksheet ActiveWorksheet( this MSExcel.Application application ) => ( application.ActiveSheet as MSExcel.Worksheet )!;

	public static T? RangeOrNull<T>( this MSExcel.Workbook workbook, string name )
	{
		MSExcel.Name? namedRange = null;
		try
		{
			namedRange =
				workbook.Names.Cast<MSExcel.Name>()
					.Where( n => n.Name == name )
					.FirstOrDefault();

			var range = namedRange?.RefersToRange;

			if ( typeof( T ) == typeof( string ) )
			{
				return (T?)range?.Text;
			}
			else
			{
				return (T?)range;
			}
		}
		catch ( Exception ex )
		{
			throw new ApplicationException(
				namedRange != null
					? $"Unable to get global named range of {name}.  Currently refers to {namedRange.RefersTo}."
					: $"Unable to get global named range of {name}.", ex );
		}
	}

	public static T? RangeOrNull<T>( this MSExcel.Worksheet worksheet, string name )
	{
		var sheetName = "Unavailable";
		MSExcel.Name? namedRange = null;
		try
		{
			sheetName = worksheet.Name;
			namedRange = worksheet.Names.Cast<MSExcel.Name>()
						.Where( n => n.Name.EndsWith( "!" + name ) )
						.FirstOrDefault();

			var range = namedRange?.RefersToRange;

			if ( typeof( T ) == typeof( string ) )
			{
				return (T?)range?.Text;
			}
			else
			{
				return (T?)range;
			}
		}
		catch ( Exception ex )
		{
			throw new ApplicationException(
				namedRange != null
					? $"Unable to get global named range of {sheetName}!{name}.  Currently refers to {namedRange.RefersTo}."
					: $"Unable to get global named range of {sheetName}!{name}.", ex );
		}
	}

	public static string GetText( this MSExcel.Range range ) => ( range.Text as string )!;

	public static MSExcel.Range GetRange( this string address, MSExcel.Worksheet worksheet )
	{
		// Start: '[Buck_MurphyOil_SE debug macro.xls]RBLMacro'!$A$3

		var addressParts = address.Split( '!' );
		// '[Buck_MurphyOil_SE debug macro.xls]RBLMacro'
		// $A$3

		if ( addressParts.Length == 1 )
		{
			return worksheet.Range[ addressParts[ 0 ] ];
		}

		var sheetName = addressParts[ 0 ].StartsWith( "'" )
			? addressParts[ 0 ][ 1..^1 ]
			: addressParts[ 0 ];
		// Current: '[Buck_MurphyOil_SE debug macro.xls]RBLMacro'

		var rangeAddress = addressParts[ 1 ];
		// Current: $A$3

		addressParts = sheetName.Split( ']' );
		// [Buck_MurphyOil_SE debug macro.xls
		// RBLMacro

		sheetName = addressParts.Last();

		if ( sheetName.EndsWith( ".csv", StringComparison.InvariantCultureIgnoreCase ) )
		{
			// Seems a csv file with one tab (the normal format) comes through as
			// address -> 'sheetname.csv'!address
			sheetName = Path.GetFileNameWithoutExtension( sheetName );
		}

		return ( ( addressParts.Length == 2 ? worksheet.Application.Workbooks[ addressParts[ 0 ][ 1.. ] ] : worksheet.Application.ActiveWorkbook ).Worksheets[ sheetName ] as MSExcel.Worksheet )!.Range[ rangeAddress ];
	}
}