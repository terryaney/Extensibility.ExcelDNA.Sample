using ExcelDna.Integration;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public class KatExcelArgumentAttribute : ExcelArgumentAttribute
{
	public string? Summary = null;
	public string? DisplayName = null;
	public Type? Type = null;
	public string? Default = null;

	public KatExcelArgumentAttribute() { }

	public KatExcelArgumentAttribute( string description )
	{
		Description = description;
	}
}
