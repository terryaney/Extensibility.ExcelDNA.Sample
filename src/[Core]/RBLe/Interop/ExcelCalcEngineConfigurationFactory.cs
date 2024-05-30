using MSExcel = Microsoft.Office.Interop.Excel;

using KAT.Camelot.RBLe.Core.Calculations;

namespace KAT.Camelot.Extensibility.Excel.AddIn.RBLe.Interop;

public class ExcelCalcEngineConfigurationFactory : CalcEngineConfigurationFactory<MSExcel.Workbook, MSExcel.Worksheet, MSExcel.Range>
{
	private readonly MSExcel.Workbook workbook;

	public ExcelCalcEngineConfigurationFactory( MSExcel.Workbook workbook ) => this.workbook = workbook;

	protected override string FileName => workbook.Name;
	protected override string Version => workbook.RangeOrNull<string>( "Version" )!;
	protected override MSExcel.Worksheet[] Worksheets => workbook.Worksheets.Cast<MSExcel.Worksheet>().ToArray();
	protected override MSExcel.Worksheet GetSheet( MSExcel.Range range ) => range.Worksheet;
	protected override string GetName( MSExcel.Worksheet sheet ) => sheet.Name;
	protected override string? RangeTextOrNull( MSExcel.Worksheet sheet, string name ) => sheet.RangeOrNull<string>( name );
	protected override MSExcel.Range GetRange( string name ) => workbook.RangeOrNull( name )!;
	protected override MSExcel.Range GetRange( MSExcel.Worksheet sheet, string name ) => sheet.Range[ name ];
	protected override MSExcel.Range Offset( MSExcel.Range range, int rowOffset, int columnOffset ) => range.Offset[ rowOffset, columnOffset ];
	protected override MSExcel.Range EndRight( MSExcel.Range range ) => range.End[ MSExcel.XlDirection.xlToRight ];
	protected override bool RangeExists( string name ) => workbook.RangeOrNull( name ) != null;
	protected override bool RangeExists( MSExcel.Worksheet sheet, string name ) => sheet.RangeOrNull( name ) != null;
	protected override string GetAddress( MSExcel.Range range ) => range.Address;
	protected override string GetText( MSExcel.Range range ) => range.GetText();
}