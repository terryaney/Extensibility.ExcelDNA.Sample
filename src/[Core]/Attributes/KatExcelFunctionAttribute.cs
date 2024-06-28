using ExcelDna.Integration;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public class KatExcelFunctionAttribute : ExcelFunctionAttribute
{
	public string? Returns = null;
	public string? Summary = null;
	public string? Remarks = null;
	public string? Exceptions = null;
	public string? Example = null;
	public bool CreateDebugFunction = false;
	
	public KatExcelFunctionAttribute() { }

	public KatExcelFunctionAttribute( string description )
	{
		Description = description;
	}
}
