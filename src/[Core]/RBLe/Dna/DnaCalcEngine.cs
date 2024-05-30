using System.Text.RegularExpressions;
using ExcelDna.Integration;
using KAT.Camelot.Extensibility.Excel.AddIn.ExcelApi;
using KAT.Camelot.RBLe.Core.Calculations;
using XLParser;

namespace KAT.Camelot.Extensibility.Excel.AddIn.RBLe.Dna;

class DnaCalcEngine : ICalcEngine<DnaWorkbook, DnaWorksheet, ExcelReference, ExcelError>, IDisposable
{
	private static readonly string[] _MacroFunctionNames = new[] { "BTRGetMacroVariable" };
	private readonly DnaWorkbook workbook;
	private DnaWorksheet? wsDnaEvaluate;
	private ExcelReference? evalReference;
	private readonly string version;
	private readonly string fileName;

	private int? macroTimeout;

	public DnaCalcEngine( string fileName ) : this( fileName, null! ) { }
	public DnaCalcEngine( string fileName, CalcEngineConfiguration configuration )
	{
		this.fileName = fileName;
		workbook = new( fileName );
		version = workbook.Version!;

		Configuration = configuration;
	}

	public CalcEngineConfiguration Configuration { get; init; } = null!;

	public string[] MacroFunctionNames => _MacroFunctionNames;
	public int? MacroTimeout => macroTimeout;
	public string FileName => fileName;
	public string Version => version;
	public ExcelError NoMatchValue => ExcelError.ExcelErrorNA;
	public int GetRow( ExcelReference cell ) => cell.RowFirst;
	
	public void GoalSeek( ExcelReference valueCell, double goal, ExcelReference changingCell ) => XlCall.Excel( XlCall.xlcGoalSeek, valueCell, goal, changingCell );

	public void SortRange( ExcelReference sortRange, SortKey<ExcelReference> key1, SortKey<ExcelReference>? key2, SortKey<ExcelReference>? key3, bool sortByColumns, bool matchCase ) =>
		sortRange.Sort( key1, key2, key3, sortByColumns, matchCase );

	public RangeDimension GetDimensions( ExcelReference range ) => new() { Columns = range.ColumnLast - range.ColumnFirst + 1, Rows = range.RowLast - range.RowFirst + 1 };
	public ExcelReference Corner( ExcelReference range, CornerType cornerType ) => range.Corner( cornerType );

	public DnaWorksheet[] Worksheets => workbook.Worksheets;
	public DnaWorksheet? GetSheet( string name ) => new( workbook.Name, name );
	public string GetName( DnaWorksheet sheet ) => sheet.Name;
	public string? RangeTextOrNull( DnaWorksheet sheet, string name ) => sheet.ReferenceOrNull<string>( name );
	public ExcelReference GetRange( DnaWorksheet sheet, string name ) => sheet.ReferenceOrNull( name )!;
	public ExcelReference GetRange( string name ) => workbook.ReferenceOrNull( name )!;
	public ExcelReference? GetRangeFromAddress( string address ) => DnaApplication.GetRangeFromAddress( address );

	public ExcelReference Offset( ExcelReference range, int rowOffset, int columnOffset ) => range.Offset( rowOffset, columnOffset );
	public bool RangeExists( string name ) => workbook.ReferenceOrNull( name ) != null;
	public string GetA1Address( ExcelReference range ) => range.GetAddress().Split( '!' ).Last();
	public string GetFullAddress( ExcelReference range )
	{
		var address = range.GetAddress();
		var pos = Math.Max( address.IndexOf( "]" ) + 1, 0 );
		return address[ pos.. ];
	}
	public string GetText( ExcelReference range ) => range.GetValue<string>() ?? "";
	public string GetFormula( ExcelReference range ) => range.GetFormula()!;
	public string GetExportValue( ExcelReference range )
	{
		var text = range.GetValue<string>();
		return text == "'" ? "" : text!;
	}
	public bool TextForced( ExcelReference range ) => !string.IsNullOrEmpty( range.PrefixCharacter() ) || range.GetValue<string>() == "'";
	public void ClearContents( ExcelReference range ) => range.ClearContents();
	public void SetValue( ExcelReference range, object value ) => range.SetValue( value );
	public object GetValue( ExcelReference range ) => range.GetValue<object>()!;

	public T[,] GetArray<T>( ExcelReference range ) => range.GetArray<T>()!;
	public T[] GetValues<T>( ExcelReference range ) => range.GetValues<T>()!;

	public void SetArray<T>( ExcelReference range, T[] array )
	{
		if ( array.Length == 0 ) return;

		var target = new ExcelReference( 
			range.RowFirst, 
			range.RowFirst + array.Length - 1, 
			range.ColumnFirst, 
			range.ColumnFirst, 
			range.SheetId 
		);

		target.SetArray( array );
	}
	public void SetArray<T>( ExcelReference range, T[,] array )
	{
		if ( array.GetUpperBound( 0 ) == -1 ) return;

		var target = new ExcelReference( 
			range.RowFirst, 
			range.RowFirst + array.GetUpperBound( 0 ), 
			range.ColumnFirst, 
			range.ColumnFirst + array.GetUpperBound( 1 ), 
			range.SheetId 
		);

		target.SetValue( array );
	}
	public void CopyAddress( ExcelReference source, ExcelReference destination ) => destination.SetValue( source );
	public void CopyRange( ExcelReference source, ExcelReference destination ) => XlCall.Excel( XlCall.xlcCopy, source, destination );
	public void TransposeAddress( ExcelReference source, ExcelReference destination )
	{
		var sourceData = source.GetArray<object>()!;

		var destinationData = new object[ sourceData.GetLength( 1 ), sourceData.GetLength( 0 ) ];
		for ( var i = 0; i < sourceData.GetLength( 0 ); i++ )
		{
			for ( var j = 0; j < sourceData.GetLength( 1 ); j++ )
			{
				destinationData[ j, i ] = sourceData[ i, j ]!;
			}
		}

		destination.SetValue( destinationData );
	}
	public void FillDown( ExcelReference range, int rows ) => Extend( range, Offset( range, rows, 0 ) ).FillDown();
	public ExcelReference End( ExcelReference range, DirectionType direction ) => range.End( direction );
	public ExcelReference Extend( ExcelReference start, ExcelReference end ) => start.Extend( end );

	public int Rows( ExcelReference table ) => table.End( DirectionType.Down ).RowFirst - table.RowFirst;
	public T EvaluateFormula<T>( ExcelReference _, string formula ) 
	{
		if ( wsDnaEvaluate == null )
		{
			XlCall.Excel( XlCall.xlcActivate, workbook.Name );
			XlCall.Excel( XlCall.xlcWorkbookInsert, (int)WorkbookInsertType.Worksheet );
			evalReference = XlCall.Excel( XlCall.xlfSelection ) as ExcelReference;
			var oldName = XlCall.Excel( XlCall.xlfGetCell, (int)GetCellType.SheetRef, evalReference );
			XlCall.Excel( XlCall.xlcWorkbookName, oldName, "_macroEvaluate" );
			wsDnaEvaluate = new DnaWorksheet( workbook.Name, "_macroEvaluate" );
			XlCall.Excel( XlCall.xlcWorkbookHide, "_macroEvaluate" );
		}

		// https://xlparser.perfectxl.nl/demo/
		var apiFormula = formula;
		var tree = ExcelFormulaParser.Parse( apiFormula );
		var references = tree.AllNodes( GrammarNames.Reference ).Where( r => new [] { GrammarNames.Prefix, GrammarNames.Cell }.Contains( r.ChildNodes[ 0 ].Type() ) ).ToArray();

		foreach( var r in references )
		{
			var count = r.ChildNodes.Count;
			var a1Prefix = count == 1 ? ":" : r.ChildNodes[ 0 ].ChildNodes[ 0 ].Token.Text;
			var a1Address = count == 1
				? r.ChildNodes[ 0 ].ChildNodes[ 0 ].Token.Text
				: r.ChildNodes[ 1 ].ChildNodes[ 0 ].Token.Text;

			// Extract the column letter and row number from the A1 address
			var r1c1Address = a1Address.Replace( "$", "" );
			var match = Regex.Match( r1c1Address, @"([A-Z]+)(\d+)" );

			if ( match.Success )
			{
				var columnLetter = match.Groups[ 1 ].Value;
				var rowNumber = int.Parse( match.Groups[ 2 ].Value );

				// Convert the column letter to a column number
				var columnNumber = 0;
				for ( int i = columnLetter.Length - 1, j = 0; i >= 0; i--, j++ )
				{
					columnNumber += ( columnLetter[ i ] - 'A' + 1 ) * (int)Math.Pow( 26, j );
				}

				// Return the address in R1C1 format
				r1c1Address = $"R{rowNumber}C{columnNumber}";
				apiFormula = apiFormula.Replace( $"{a1Prefix}{a1Address}", $"{a1Prefix}{r1c1Address}" );
			}
		}

		// apiFormula = (string)XlCall.Excel( XlCall.xlfFormulaConvert, formula, true, false );
		XlCall.Excel( XlCall.xlcFormula, apiFormula, evalReference );
		var dnaValue = evalReference!.GetValue();
		
		try
		{
			if ( dnaValue is ExcelError r )
			{
				throw new ApplicationException( r == ExcelError.ExcelErrorRef
					? $"Processing Macro variables in formula << {formula} >> failed.  Make sure all cell references have Sheet! prefix in the formula, even if cells are located on RBLMacro tab."
					: $"Processing Macro variables in formula << {formula} >> failed returning {r}."
				 );
			}

			return (T)dnaValue;
		}
		catch ( Exception ex ) when ( ex is not ApplicationException )
		{
			throw new ApplicationException( $"Unable to evaluate formula '{formula}'.  Text result is {dnaValue?.ToString()}.", ex );
		}
	}

	public void Calculate() => XlCall.Excel( XlCall.xlcCalculateNow );

	private CalculationType CalculationType => workbook.CalculationType;

	public MacroCalculationType GetCalculationMode( string calculationKey )
	{
		if ( calculationKey == "Application" )
		{
			return CalculationType == CalculationType.Automatic
				? MacroCalculationType.Automatic
				: MacroCalculationType.Manual;
		}
		else
		{
			return MacroCalculationType.Automatic;
			// throw new NotImplementedException( "Can't get calculation mode on just a worksheet." );
		}
	}

	public void EnsureManualCalculation()
	{
		if ( CalculationType != CalculationType.Manual )
		{
			XlCall.Excel( XlCall.xlcCalculation, (int)CalculationType.Manual );
		}
	}
	public void EnsureAutomaticCalculation()
	{
		if ( CalculationType != CalculationType.Automatic )
		{
			XlCall.Excel( XlCall.xlcCalculation, (int)CalculationType.Automatic );
			Calculate();
		}
	}

	public bool SetCalculationMode( string calculationKey, MacroCalculationType calculationType )
	{
		// See SpreadsheetGearCalcEngine for documentation.
		if ( calculationKey == "Application" )
		{
			XlCall.Excel( XlCall.xlcCalculation, calculationType == MacroCalculationType.Automatic ? (int)CalculationType.Automatic : (int)CalculationType.Manual );

			return true;
		}

		return calculationKey == "_CalculationOnDemand";
	}

	public bool ProcessMacroAction( MacroInstruction<ExcelReference> action, Dictionary<string, MacroCalculationType> calculationModes, Action<MacroInstruction<ExcelReference>?, string?> traceMacroAction )
	{
		var actionName = action.Action;

		if ( string.Compare( actionName, "CalculationOnDemand", true ) == 0 )
		{
			throw new NotImplementedException( "CalculationOnDemand is not supported.  See documentation in SpreadsheetGearCalcEngine for possible work arounds." );
		}
		// Do nothing on RBLe, we don't support bumping this up on our servers...
		else if ( string.Compare( actionName, "SetTimeout", true ) == 0 )
		{
			macroTimeout = (int)(double)action.Value!;
			traceMacroAction( action, string.Format( "{0:hh:mm:ss.ff} RBLe Macro: {1}/{2} - SetTimeout to {3}.", DateTime.Now, actionName, action.Address, action.Value ) );
			return true;
		}

		return false;
	}

	bool? calculationOnDemand;
	public void StartProcessMacros() 
	{
		calculationOnDemand = null;
	}

	public void FinishProcessMacros() 
	{
		if ( calculationOnDemand != null )
		{
			throw new NotImplementedException( "CalculationOnDemand is not supported.  See documentation in SpreadsheetGearCalcEngine for possible work arounds." );
			// application.CalculationOnDemand = calculationOnDemand.Value;
		}
	}

	public void Dispose()
	{
		if ( wsDnaEvaluate != null )
		{
			XlCall.Excel( XlCall.xlcWorkbookDelete, wsDnaEvaluate.Name );
		}
	}
}
