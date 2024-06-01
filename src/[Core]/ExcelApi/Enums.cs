namespace KAT.Camelot.Extensibility.Excel.AddIn.ExcelApi;

enum GetCellType
{
	Formula = 6,
	PrefixCharacter = 52,
	Text = 53,
	SheetRef = 62,
	WorkbookRef = 66
}

enum CalculationType
{
	Automatic = 1,
	Manual = 3
}

enum WorkbookInsertType
{
	Worksheet = 1
}

enum GetNameInfoType
{
	Definition = 1,
	Scope = 2 // true = sheet, false = workbook
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

enum SortOrientationType
{
	Rows = 1,
	Columns = 2
}

enum SortOrderType
{
	Ascending = 1,
	Descending = 2
}

enum SortHeaderType
{
	Guess = 0,
	Yes = 1,
	No = 2
}

enum SortDataType
{
	Values = 1,
	Data = 2
}