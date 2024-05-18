using ExcelDna.Integration;
using KAT.Camelot.RBLe.Core;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

// https://excel-dna.net/docs/guides-basic/excel-programming-interfaces/excel-c-api
// https://learn.microsoft.com/en-us/office/client-developer/excel/welcome-to-the-excel-software-development-kit
public static partial class ExcelApi
{
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

			// remember the current active cell
			var current = new
			{
				Selection = XlCall.Excel( XlCall.xlfSelection ),
				Cell = XlCall.Excel( XlCall.xlfActiveCell )
			};

			if ( reference != null )
			{
				// select caller worksheet containing range caller desires to be active
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
				// NOTE: SelectionHelper seemed to use  (https://github.com/Excel-DNA/Samples/blob/9b1b8b6c749216147352c155766556868eaae949/Archive/Async/AsyncMacros.dna#L189)
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

	public static ExcelReference? GetReferenceOrNull( this string address, string sheetName ) => $"'{sheetName}'!{address}".GetReferenceOrNull();
	public static ExcelReference? GetReferenceOrNull( this string address ) => TryGetWorkbookReference( null, address );

	public static ExcelReference GetReference( this string address, string sheetName ) => $"'{sheetName}'!{address}".GetReference();
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
	}
}