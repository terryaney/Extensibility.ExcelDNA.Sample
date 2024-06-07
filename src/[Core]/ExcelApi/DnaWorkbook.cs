using ExcelDna.Integration;

#pragma warning disable CA1822 // Doesn't access instance data warning...

namespace KAT.Camelot.Extensibility.Excel.AddIn.ExcelApi;

class DnaWorkbook
{
	public readonly string Name;
	public readonly string? Version;

	public DnaWorkbook() : this( DnaApplication.ActiveWorkbookName() ) { }
	public DnaWorkbook( string name ) 
	{
		Name = name;
		Version = RangeOrNull( "Version" )?.GetValue<string>();
	} 
	public bool IsSaved => (bool)XlCall.Excel( XlCall.xlfGetWorkbook, (int)GetWorkbookType.IsSaved, Name );

	public CalculationType CalculationType => (CalculationType)(int)(double)XlCall.Excel( XlCall.xlfGetDocument, (int)GetDocumentType.CalculationMode );
	
	public DnaWorksheet[] Worksheets
	{
		get
		{
			var sheets = (object[,])XlCall.Excel( XlCall.xlfGetWorkbook, (int)GetWorkbookType.SheetNames, Name );
			var count = sheets.GetLength( 1 );

			return Enumerable.Range( 0, count )
				.Select( i => new DnaWorksheet( Name, ( (string)sheets[ 0, i ] ).Split( ']' ).Last() ) )
				.ToArray();
		}
	}

	public ExcelReference? RangeOrNull( string address )
	{
		// C API scope issue: https://stackoverflow.com/questions/78551625/excel-c-api-via-exceldna-and-named-range-scopes

		var sheetName = (string)XlCall.Excel( XlCall.xlfGetWorkbook, (int)GetWorkbookType.ActiveSheet, Name );
		var scopeResult = XlCall.Excel( XlCall.xlfGetName, $"'[{Name}]{sheetName}'!{address}", (int)GetNameInfoType.Scope );

		if ( scopeResult is ExcelError )
		{
			return null;
		}

		if ( (bool)scopeResult )
		{
			// Sheet scope...
			var otherSheet = Worksheets.FirstOrDefault( w => w.Name != sheetName );
			if ( otherSheet == null )
			{
				return null; // Unable to get a different sheet so I can get 'workbook scope' range....
			}

			XlCall.Excel( XlCall.xlcWorkbookActivate, otherSheet.Name );

			try
			{
				return RangeOrNull( address );				
			}
			finally
			{
				XlCall.Excel( XlCall.xlcWorkbookActivate, sheetName );
			}
		}

		return XlCall.Excel( XlCall.xlfEvaluate, $"='[{Name}]{sheetName}'!{address}" ) as ExcelReference;
	}

	public void Activate() => XlCall.Excel( XlCall.xlcActivate, Name );
}