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

	public ExcelReference? RangeOrNull( string nameOrAddress )
	{
		// C API scope issue: https://stackoverflow.com/questions/78551625/excel-c-api-via-exceldna-and-named-range-scopes

		var sheetName = (string)XlCall.Excel( XlCall.xlfGetWorkbook, (int)GetWorkbookType.ActiveSheet, WorkbookName );
		
		if ( sheetName != Name )
		{
			XlCall.Excel( XlCall.xlcWorkbookActivate, Name );
		}

		try
		{
			var scopeResult = XlCall.Excel( XlCall.xlfGetName, $"'[{WorkbookName}]{Name}'!{nameOrAddress}", (int)GetNameInfoType.Scope );
			if ( scopeResult is bool b && !b )
			{
				// Found named range but scoped to Workbook...
				return null;
			}

			return XlCall.Excel( XlCall.xlfEvaluate, $"'[{WorkbookName}]{Name}'!{nameOrAddress}" ) as ExcelReference;
		}
		finally
		{
			if ( sheetName != Name )
			{
				XlCall.Excel( XlCall.xlcWorkbookActivate, sheetName );
			}
		}			
	}
}