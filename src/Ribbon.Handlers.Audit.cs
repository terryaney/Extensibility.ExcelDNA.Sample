using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using KAT.Camelot.Domain.Extensions;
using KAT.Camelot.RBLe.Core.Calculations;
using XLParser;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public partial class Ribbon
{
	public void Audit_ShowDependencies( IRibbonControl _ )
	{
		foreach ( MSExcel.Range cell in ( application.Selection as MSExcel.Range )! )
		{
			cell.ShowDependents();
		}
	}

	public void Audit_HideDependencies( IRibbonControl _ )
	{
		foreach ( MSExcel.Range cell in ( application.Selection as MSExcel.Range )! )
		{
			cell.ShowDependents( true );
		}
	}

	public void Audit_ShowCellsWithEmptyDependencies( IRibbonControl _ )
	{
		var selection = ( application.Selection as MSExcel.Range )!;
		selection.Style = "Normal";

		var selectionRef = selection.GetReference();
		var firstCell = selectionRef.Corner( CornerType.UpperLeft );

		for ( var row = 0; row <= selectionRef.RowLast - selectionRef.RowFirst; row++ )
		{
			for ( var col = 0; col <= selectionRef.ColumnLast - selectionRef.ColumnFirst; col++ )
			{
				var cell = firstCell.Offset( row, col );
				var cellFormula = cell.GetFormula();

				if ( !string.IsNullOrEmpty( cellFormula ) && cellFormula.StartsWith( "=" ) )
				{
					var tree = ExcelFormulaParser.Parse( cellFormula );
					var references =
						tree.AllNodes()
							.Where( n => n.Term.Name == "NamedRange" || n.Term.Name == "Cell" )
							.Select( r => r.ChildNodes[ 0 ].Token.Text );

					foreach ( var r in references )
					{
						var reference = r.GetReference();
						var data = reference.GetValue();
						var dataValues = data as object[,];

						if ( dataValues?.Contains( ExcelEmpty.Value, false ) ?? Equals( data, ExcelEmpty.Value ) )
						{
							cell.GetRange().Style = "Bad";
							break;
						}
					}
				}
			}
		}
	}

	public void Audit_SearchLocalCalcEngines( IRibbonControl _ )
	{
		
	}

	public void Audit_CalcEngineTabs( IRibbonControl control )
	{
		MessageBox.Show( "// TODO: Process " + control.Id );
	}
}