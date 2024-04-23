using ExcelDna.Integration.CustomUI;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public void Navigation_NavigateToTable( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}
	
	public void Navigation_GoToInputs( IRibbonControl _ ) => GotoInputNamedRange( "StartData" );
	public void Navigation_GoToInputData( IRibbonControl _ ) => GotoInputNamedRange( "xDSDataFields" );
	public void Navigation_GoToCalculationInputs( IRibbonControl _ ) => GotoInputNamedRange( "CalculationInputs" );
	public void Navigation_GoToFrameworkInputs( IRibbonControl _ ) => GotoInputNamedRange( "FrameworkInputs" );
	public void Navigation_GoToInputTables( IRibbonControl _ ) => GotoInputNamedRange( "StartTables" );

	public void Navigation_GoToBTRCellAddress( IRibbonControl _ )
	{
		var formula = "[Unavailable]";
		var address = "[Unavailable]";
		try
		{
			var selection = application.Selection as MSExcel.Range;
			formula = selection!.Formula as string;
			
			if ( formula!.Contains( "BTRCellAddress" ) )
			{
				address = selection.Text as string;
				var range = address!.GetRange( application.ActiveWorksheet() );
				range.Worksheet.Activate();
				range.Activate();
			}
		}
		catch ( Exception ex )
		{
			throw new ApplicationException( $"Unable to go to BTRCellAddress selected.\r\n\r\nFormula: {formula}\r\nAddress: {address}", ex );
		}
	}

	public void Navigation_BackToRBLeMacro( IRibbonControl _ ) => GotoNamedRange( "RBLeMacro", false );

	private void GotoInputNamedRange( string name )
	{
		var inputTab =
			application
				.ActiveWorkbook
				.Sheets
				.Cast<MSExcel.Worksheet>()
				.FirstOrDefault( w => w.RangeOrNull<string>( "SheetType" ) == "Input" );

		if ( inputTab != null )
		{
			GotoNamedRange( $"{inputTab.Name}!{name}", true );
		}
	}

	private void GotoNamedRange( string name, bool activate )
	{
		var range =
			application.ActiveWorkbook.Names.Cast<MSExcel.Name>().FirstOrDefault( n => n.Name == name )?.RefersToRange ??
			application.ActiveWorkbook.Names.Cast<MSExcel.Name>().FirstOrDefault( n => n.Name == name.Split( '!' ).Last() )?.RefersToRange; // Incase they didn't scope sheet name properly, remove sheet and try

		if ( range != null )
		{
			range.Worksheet.Activate();
			if ( activate )
			{
				range.Select();
			}
		}
	}
}