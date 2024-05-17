namespace KAT.Camelot.Extensibility.Excel.AddIn;

enum GetCellType
{
	Formula = 6,
	Text = 53,
	SheetRef = 62,
	WorkbookRef = 66
}

enum GetWorkbookType
{
	SheetNames = 1,
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
