using ExcelDna.Integration;
using KAT.Camelot.RBLe.Core;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public static partial class ExcelApi
{
	enum GetCellType
	{
		Formula = 6,
		Text = 53,
		SheetRef = 62
	}

	enum GetWorkbookType
	{
		IsSaved = 24,
		ActiveSheet = 38
	}

	enum GetDocumentType
	{
		ActiveWorkbookPath = 2,
		CalculationMode = 14,
		ActiveSheet = 76, // in the form [Book1]Sheet1
		ActiveWorkbook = 88
	}

	enum GetWorkspaceType
	{
		ScreenUpdating = 40
	}

	public static ExcelReference? GetCaller()
	{
		var caller = XlCall.Excel( XlCall.xlfCaller );
		return caller is ExcelReference reference ? reference : null;
	}

	public static string ActiveWorkbookName() => (string)XlCall.Excel( XlCall.xlfGetDocument, (int)GetDocumentType.ActiveWorkbook );
	
	public static bool ScreenUpdating
	{
		set { XlCall.Excel( XlCall.xlcEcho, value ); }
		get { return (bool)XlCall.Excel( XlCall.xlfGetWorkspace, (int)GetWorkspaceType.ScreenUpdating ); }
	}

	public static ExcelReference Selection => ( XlCall.Excel( XlCall.xlfSelection ) as ExcelReference )!;

	public static void RestoreSelection( this ExcelReference reference, Action action )
	{
		// https://groups.google.com/d/msg/exceldna/h3SqSA8DkPc/X0uxH4pUBgAJ - read comment about his SelectionHelper

		var updating = ScreenUpdating;

		try
		{
			if ( updating ) ScreenUpdating = false;

			//remember the current active cell
			var current = new
			{
				Selection = XlCall.Excel( XlCall.xlfSelection ),
				Cell = XlCall.Excel( XlCall.xlfActiveCell )
			};

			if ( reference != null )
			{
				//select caller worksheet containing range caller desires to be active
				// (need to do this before reading selection/cell on this sheet)
				var rangeSheet = (string)XlCall.Excel( XlCall.xlSheetNm, reference );
				XlCall.Excel( XlCall.xlcWorkbookSelect, new object[] { rangeSheet } );
			}

			// record selection and active cell on the sheet we want to select
			var sheetCurrent = reference != null
				? new
				{
					Selection = XlCall.Excel( XlCall.xlfSelection ),
					Cell = XlCall.Excel( XlCall.xlfActiveCell ),
					OriginalSheet = (string)XlCall.Excel( XlCall.xlSheetNm, current.Selection )
				}
				: null;

			if ( reference != null )
			{
				// Select the range caller desires to be active...
				XlCall.Excel( XlCall.xlcSelect, reference );
			}

			action();

			// Now restore everything...

			if ( reference != null )
			{
				// Reset the selection on the target sheet
				XlCall.Excel( XlCall.xlcSelect, sheetCurrent!.Selection, sheetCurrent.Cell );

				// Reset the sheet originally selected
				XlCall.Excel( XlCall.xlcWorkbookSelect, new object[] { sheetCurrent.OriginalSheet } );
			}

			// Reset the selection in the active sheet (some bugs make this change sometimes too)
			XlCall.Excel( XlCall.xlcSelect, current.Selection, current.Cell );
		}
		finally
		{
			if ( updating ) ScreenUpdating = true;
		}
	}

	public static ExcelReference GetReference( this string address ) => GetWorkbookReference( null, address );

	private static ExcelReference GetWorkbookReference( string? workbookName, string address )
	{
		return
			TryGetWorkbookReference( workbookName, address ) ?? 
			throw new ArgumentOutOfRangeException( 
				nameof( address ), 
				string.Compare( workbookName, Constants.FileNames.Helpers, true ) == 0
					? $"The address provided, {workbookName}.{address}, is not a valid address.  Please make sure you also have the '{Constants.FileNames.Helpers}' worksheet open as well."
					: $"The address provided, {workbookName}.{address}, is not a valid address."
			);
	}

	private static ExcelReference TryGetWorkbookReference( string? workbookName, string address )
	{
		var excelAddress = address.GetExcelAddress();

		var hasSheet = !string.IsNullOrEmpty( excelAddress.Sheet ) && excelAddress.Sheet != ( workbookName ?? ActiveWorkbookName() ); /* global range if so */
		if ( !hasSheet )
		{
			// Assuming if no sheet provided, they mean the 'active' sheet
			excelAddress = new()
			{
				Address = excelAddress.Address,
				Workbook = excelAddress.Workbook,
				Sheet = !string.IsNullOrEmpty( workbookName )
					? (string)XlCall.Excel( XlCall.xlfGetWorkbook, (int)GetWorkbookType.ActiveSheet, workbookName )
					: ( (string)XlCall.Excel( XlCall.xlfGetDocument, (int)GetDocumentType.ActiveSheet ) ).Split( ']' ).Last()
			};
		}

		var addressWorkbook = excelAddress.Workbook ?? workbookName;
		var workbookNameSyntax = !string.IsNullOrEmpty( addressWorkbook ) ? $"[{addressWorkbook}]" : null;

		var fullAddress = $"='{workbookNameSyntax}{excelAddress.Sheet}'!{excelAddress.Address}";
		var reference = XlCall.Excel( XlCall.xlfEvaluate, fullAddress ) as ExcelReference;

		return reference!;

		// http://www.technicana.com/xlftextref.pdf
		// Could try to get rid of all Interop objects during macro processing and attempt to use item below
		/*
		XlCall.Excel(XlCall.xlfTextref, referenceString, a1)

		referenceString is a reference in string format.

		a1 is a boolean value specifying the form the reference is in.  If a1 is TRUE, referenceString should be in A1-style.  If a1 is FALSE, referenceString must be in R1C1 format.  The default is FALSE.
		*/
	}
}