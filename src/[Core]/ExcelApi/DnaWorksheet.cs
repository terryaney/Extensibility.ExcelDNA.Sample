using ExcelDna.Integration;

namespace KAT.Camelot.Extensibility.Excel.AddIn.ExcelApi;

class DnaWorksheet
{
	public readonly string Name;
	public readonly string WorkbookName;

	public DnaWorksheet( string workbookName, string name )
	{
		WorkbookName = workbookName;
		Name = name;
	}

	public ExcelReference? ReferenceOrNull( string address ) => XlCall.Excel( XlCall.xlfEvaluate, $"='[{WorkbookName}]{Name}'!{address}" ) as ExcelReference;
	public T? ReferenceOrNull<T>( string name ) => ReferenceOrNull( name )!.GetValue<T>();
}