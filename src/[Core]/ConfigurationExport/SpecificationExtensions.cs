using KAT.Camelot.Domain.Extensions;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn.ConfigurationExport;

static class SpecificationExtensions
{
	public static string? GetValue<TEnum>( this Dictionary<TEnum, ColumnDefinition> configuration, TEnum column, MSExcel.Range range ) where TEnum : struct, Enum
	{
		if ( !configuration.TryGetValue( column, out var colConfig ) ) return null;

		return range.Offset[ 0, colConfig.Offset ].GetText().EnsureNonBlank();
	}

	public static string? GetValue( this Dictionary<string, ColumnDefinition> configuration, string column, MSExcel.Range range )
	{
		if ( !configuration.TryGetValue( column, out var colConfig ) ) return null;

		return range.Offset[ 0, colConfig.Offset ].GetText().EnsureNonBlank();
	}
}