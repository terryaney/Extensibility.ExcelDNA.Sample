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
}