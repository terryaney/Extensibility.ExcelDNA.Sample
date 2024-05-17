using KAT.Camelot.RBLe.Core;
using KAT.Camelot.RBLe.Core.Calculations;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn.RBLe.Interop;

public class ExcelCalcEngine : ICalcEngine<MSExcel.Workbook, MSExcel.Worksheet, MSExcel.Range, MSExcel.XlCVError>, IDisposable
{
	private readonly bool isSaved;
	private readonly MSExcel.Worksheet wsEvaluate;
	private readonly MSExcel.Application application;
	private readonly string version;
	private readonly MSExcel.Workbook workbook;
	private int? macroTimeout;

	public ExcelCalcEngine( MSExcel.Workbook workbook ) : this( workbook, null! ) { }
	public ExcelCalcEngine( MSExcel.Workbook workbook, CalcEngineConfiguration configuration )
	{
		application = workbook.Application;
		version = workbook.RangeOrNull<string>( "Version" )!;
		this.workbook = workbook;
		isSaved = workbook.Saved;

		application.Calculation = MSExcel.XlCalculation.xlCalculationManual;
		Configuration = configuration;

		wsEvaluate = ( application.Worksheets.Add() as MSExcel.Worksheet )!;
		wsEvaluate.Visible = MSExcel.XlSheetVisibility.xlSheetHidden;
	}

	public CalcEngineConfiguration Configuration { get; init; } = null!;

	public int? MacroTimeout => macroTimeout;
	public string FileName => workbook.FullName;
	public string Version => version;
	public MSExcel.XlCVError NoMatchValue => MSExcel.XlCVError.xlErrNA;
	public int GetRow( MSExcel.Range cell ) => cell.Row;
	public void GoalSeek( MSExcel.Range valueCell, double goal, MSExcel.Range changingCell ) => valueCell.GoalSeek( goal, changingCell );
	public void SortRange( MSExcel.Range sortRange, SortKey<MSExcel.Range> key1, SortKey<MSExcel.Range>? key2, SortKey<MSExcel.Range>? key3, bool sortByColumns, bool matchCase )
	{
		sortRange.Sort(
			key1.Key,
			key1.IsAscending ? MSExcel.XlSortOrder.xlAscending : MSExcel.XlSortOrder.xlDescending,
			key2 != null ? (object)key2.Key : Type.Missing,
			Type.Missing,
			key2?.IsAscending ?? true ? MSExcel.XlSortOrder.xlAscending : MSExcel.XlSortOrder.xlDescending,
			key3 != null ? (object)key3.Key : Type.Missing,
			key3?.IsAscending ?? true ? MSExcel.XlSortOrder.xlAscending : MSExcel.XlSortOrder.xlDescending,
			MSExcel.XlYesNoGuess.xlNo,
			Type.Missing,
			matchCase,
			sortByColumns ? MSExcel.XlSortOrientation.xlSortColumns : MSExcel.XlSortOrientation.xlSortRows,
			MSExcel.XlSortMethod.xlPinYin,
			key1.IsTextAsNumbers ? MSExcel.XlSortDataOption.xlSortTextAsNumbers : MSExcel.XlSortDataOption.xlSortNormal,
			key2?.IsTextAsNumbers ?? false ? MSExcel.XlSortDataOption.xlSortTextAsNumbers : MSExcel.XlSortDataOption.xlSortNormal,
			key3?.IsTextAsNumbers ?? false ? MSExcel.XlSortDataOption.xlSortTextAsNumbers : MSExcel.XlSortDataOption.xlSortNormal
		);
	}

	public RangeDimension GetDimensions( MSExcel.Range range ) => new() { Columns = range.Columns.Count, Rows = range.Rows.Count };
	public MSExcel.Range Corner( MSExcel.Range range, CornerType cornerType )
	{
		var row = cornerType == CornerType.UpperLeft || cornerType == CornerType.UpperRight
			? range.Row
			: range.Row + range.Rows.Count - 1;

		var column = cornerType == CornerType.UpperLeft || cornerType == CornerType.LowerLeft
			? range.Column
			: range.Column + range.Columns.Count - 1;

		return ( range.Worksheet.Cells[ row, column ] as MSExcel.Range )!;
	}

	public MSExcel.Worksheet[] Worksheets => workbook.Worksheets.Cast<MSExcel.Worksheet>().ToArray();
	public MSExcel.Worksheet GetSheet( MSExcel.Range range ) => range.Worksheet;
	public MSExcel.Worksheet? GetSheet( string name ) => Worksheets.FirstOrDefault( w => w.Name == name );
	public string GetName( MSExcel.Worksheet sheet ) => sheet.Name;
	public string? RangeTextOrNull( string name ) => workbook.RangeOrNull<string>( name );
	public string? RangeTextOrNull( MSExcel.Worksheet sheet, string name ) => sheet.RangeOrNull<string>( name );
	public MSExcel.Range GetRange( MSExcel.Worksheet sheet, string name ) => sheet.Range[ name ];
	public MSExcel.Range GetRange( string name ) => workbook.RangeOrNull( name )!;
	public MSExcel.Range? GetRangeFromAddress( string address )
	{
		try
		{
			var excelAddress = address.GetExcelAddress();
			var range = string.IsNullOrEmpty( excelAddress.Sheet ) || excelAddress.Sheet == workbook.Name /* global range if so */
				? workbook.RangeOrNull( excelAddress.Address )
				: null;

			if ( range != null )
			{
				return range;
			}

			if ( string.IsNullOrEmpty( excelAddress.Sheet ) )
			{
				// No named range matching
				return null;
			}

			return workbook.GetWorksheet( excelAddress.Sheet )!.RangeOrNull( excelAddress.Address );
		}
		catch ( Exception ex )
		{
			throw new ArgumentOutOfRangeException( $"Unable to convert address '{address}' into a Range object.", ex );
		}
	}

	public MSExcel.Range Offset( MSExcel.Range range, int rowOffset, int columnOffset ) => range.Offset[ rowOffset, columnOffset ];
	public MSExcel.Range EndRight( MSExcel.Range range ) => range.End[ MSExcel.XlDirection.xlToRight ];
	public bool RangeExists( MSExcel.Worksheet sheet, string name ) => sheet.RangeOrNull( name ) != null;
	public bool RangeExists( string name ) => workbook.RangeOrNull( name ) != null;
	public string GetAddress( MSExcel.Range range ) => range.Address;
	public string GetFullAddress( MSExcel.Range range ) => $"{range.Worksheet.Name}!{range.Address}";
	public string GetText( MSExcel.Range range ) => range.GetText();
	public string GetFormula( MSExcel.Range range ) => (string)range.Formula!;
	public string GetExportValue( MSExcel.Range range )
	{
		var text = range.GetText();
		return text == "'" ? "" : text;
	}
	public bool TextForced( MSExcel.Range range ) => !string.IsNullOrEmpty( (string?)range.PrefixCharacter ) || range.GetText() == "'";
	public void ClearContents( MSExcel.Range range ) => range.ClearContents();
	public void SetValue( MSExcel.Range range, string value ) => range.Value = value;
	public void SetValue( MSExcel.Range range, object value )
	{
		if ( value is string v )
		{
			range.Formula = v;
		}
		else
		{
			range.Value = value;
		}
	}
	public object GetValue( MSExcel.Range range )
	{
		try
		{
			return range.Value;
		}
		catch ( Exception ex )
		{
			throw new ApplicationException( $"Unable to get cell.Value for {range.Address}.  Formula: {range.Formula}", ex );
		}
	}
	public T[,] GetArray<T>( MSExcel.Range range ) => range.GetArray<T>();
	public T[] GetValues<T>( MSExcel.Range range ) => range.GetValues<T>();

	public void SetArray<T>( MSExcel.Range range, T[] array )
	{
		if ( array.Length == 0 ) return;
		var target = Extend( range, range.Offset[ array.Length - 1, 0 ] );
		target.SetArray( array );
	}
	public void SetArray<T>( MSExcel.Range range, T[,] array )
	{
		if ( array.GetUpperBound( 0 ) == -1 ) return;

		var target = Extend( range, range.Offset[ array.GetUpperBound( 0 ), array.GetUpperBound( 1 ) ] );

		// SpreadsheetGear requires it to be object[,]
		var oArray = new object[ array.GetLength( 0 ), array.GetLength( 1 ) ];
		Array.Copy( array, oArray, array.Length );

		target.Value = oArray;
	}
	public void CopyAddress( MSExcel.Range source, MSExcel.Range destination ) => destination.Value2 = source.Value2;
	public void CopyRange( MSExcel.Range source, MSExcel.Range destination ) => source.Copy( destination );
	public void TransposeAddress( MSExcel.Range source, MSExcel.Range destination )
	{
		var sourceRange = (object[,])application.WorksheetFunction.Transpose( source.Value2 );
		destination.Value2 = sourceRange;
	}
	public void FillDown( MSExcel.Range range, int rows ) => Extend( range, Offset( range, rows, 0 ) ).FillDown();
	public MSExcel.Range End( MSExcel.Range range, DirectionType direction )
	{
		var value =
			direction == DirectionType.Down ? range.Offset[ 1, 0 ].GetText() + range.Offset[ 1, 0 ].GetFormula() :
			direction == DirectionType.ToRight ? range.Offset[ 0, 1 ].GetText() + range.Offset[ 0, 1 ].GetFormula() :
			direction == DirectionType.ToLeft ? range.Offset[ 0, -1 ].GetText() + range.Offset[ 0, -1 ].GetFormula() :
			  /* DirectionType.Up */              range.Offset[ -1, 0 ].GetText() + range.Offset[ -1, 0 ].GetFormula();

		var isEmpty = string.IsNullOrEmpty( value );

		if ( isEmpty )
		{
			return range;
		}
		else
		{
			return
				direction == DirectionType.Down ? range.End[ MSExcel.XlDirection.xlDown ] :
				direction == DirectionType.ToRight ? range.End[ MSExcel.XlDirection.xlToRight ] :
				direction == DirectionType.ToLeft ? range.End[ MSExcel.XlDirection.xlToLeft ] :
				  /* DirectionType.Up */              range.End[ MSExcel.XlDirection.xlUp ];
		}
	}

	public MSExcel.Range Extend( MSExcel.Range start, MSExcel.Range end )
	{
		return start.Worksheet.Range[
			start.Worksheet.Range[ Math.Min( start.Row, end.Row ), Math.Min( start.Column, end.Column ) ],
			start.Worksheet.Range[ Math.Max( start.Row + start.Rows.Count - 1, end.Row + end.Rows.Count - 1 ), Math.Max( start.Column + start.Columns.Count - 1, end.Column + end.Columns.Count - 1 ) ]
		];
	}

	public int Rows( MSExcel.Range table ) => table.End[ MSExcel.XlDirection.xlDown ].Row - table.Row;
	public T EvalulateFormula<T>( MSExcel.Range _, string formula ) 
	{
		// https://microsoft.public.excel.sdk.narkive.com/W7118afY/strange-behaviour-of-evaluate-xlfevaluate
		wsEvaluate.Range[ "A1" ].Formula = formula;
		var value = wsEvaluate.Range[ "A1" ].Value;

		try
		{
			return (T)value;
		}
		catch ( Exception ex )
		{
			throw new ApplicationException( $"Unable to evaluate formula '{formula}'.  Text result is {wsEvaluate.Range[ "A1" ].Text}.", ex );
		}
	}

	public void Calculate() => application.Calculate();

	public MacroCalculationType GetCalculationMode( string calculationKey )
	{
		if ( calculationKey == "Application" )
		{
			return application.Calculation == MSExcel.XlCalculation.xlCalculationAutomatic
				? MacroCalculationType.Automatic
				: MacroCalculationType.Manual;
		}
		else
		{
			return ( application.Worksheets[ calculationKey ] as MSExcel.Worksheet )!.EnableCalculation
				? MacroCalculationType.Automatic
				: MacroCalculationType.Manual;
		}
	}

	public void EnsureManualCalculation()
	{
		if ( application.Calculation != MSExcel.XlCalculation.xlCalculationManual )
		{
			application.Calculation = MSExcel.XlCalculation.xlCalculationManual;
		}
	}
	public void EnsureAutomaticCalculation()
	{
		application.Calculation = MSExcel.XlCalculation.xlCalculationAutomatic;
		application.Calculate();
	}

	public bool SetCalculationMode( string calculationKey, MacroCalculationType calculationType )
	{
		if ( calculationKey == "Application" )
		{
			application.Calculation = ( calculationType == MacroCalculationType.Automatic ) ? MSExcel.XlCalculation.xlCalculationAutomatic : MSExcel.XlCalculation.xlCalculationManual;
			return true;
		}
		else if ( calculationKey != "_CalculationOnDemand" )
		{
			( application.Worksheets[ calculationKey ] as MSExcel.Worksheet )!.EnableCalculation = ( calculationType == MacroCalculationType.Automatic );
			return true;
		}

		return false;
	}

	public bool ProcessMacroAction( MacroInstruction<MSExcel.Range> action, Dictionary<string, MacroCalculationType> calculationModes, Action<MacroInstruction<MSExcel.Range>?, string?> traceMacroAction )
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
		if ( wsEvaluate != null )
		{
			application.DisplayAlerts = false;
			wsEvaluate.Delete();
			application.DisplayAlerts = true;
		}
		workbook.Saved = isSaved;
	}
}
