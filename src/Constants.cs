namespace KAT.Extensibility.Excel.AddIn;

public static class Constants
{
	public static class FileNames
	{
		public const string GlobalTables = "MadHatter_GlobalTables.xls";
		public const string RTCData = "RTCData.xlsx";
	}

	public static class CalcEngines
	{
		public static string InputSheetType = "Input";
		public static string[] PreviewSheetTypes => new[] { "Result", "ResultXml", "FolderItem" };
		public static string[] ResultSheetTypes = new[] { "Update", "ReportData" }.Concat( PreviewSheetTypes ).ToArray();
		public static string[] SheetTypes = new[] { InputSheetType }.Concat( ResultSheetTypes ).ToArray();
		public static string[] GlobalTablesSheetTypes = new[] { "Rate Tables", "Global Rate Tables", "Global Lookup Tables" };
	}
}