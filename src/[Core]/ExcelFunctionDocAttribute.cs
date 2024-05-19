using ExcelDna.Integration;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public class ExcelFunctionDocAttribute : ExcelFunctionAttribute
{
	public string? Returns = null;
	public string? Summary = null;
	public string? Remarks = null;
	public string? Example = null;

	public ExcelFunctionDocAttribute() { }

	public ExcelFunctionDocAttribute( string description )
	{
		Description = description;
	}
}
