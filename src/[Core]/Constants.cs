namespace KAT.Camelot.Extensibility.Excel.AddIn;

public static class Constants
{
	public static class FileNames
	{
		public const string GlobalTables = "MadHatter_GlobalTables.xls";
		public const string Helpers = "MadHatter_CalculationHelpers.xlsm";
		public const string RTCData = "RTCData.xlsx";
	}

	public static class SpecSheet
	{
		public static bool IsGlobalTablesSheet( string? sheetType ) => !string.IsNullOrEmpty( sheetType ) && new[] { SheetTypes.ClientRateTables, SheetTypes.GlobalLookupTables, SheetTypes.GlobalRateTables }.Contains( sheetType );

		public static class TabNames
		{
			public const string HistoricalData = "Historical Data";
			public const string CodeTables = "Code Tables";
			public const string Localization = "Localization";

			public const string DataLookupTables = "Data Lookup Tables";
			public const string RateTables = "Rate Tables";
		}

		public static class RangeNames
		{
			public const string SheetType = "SheetType";
			public const string SheetVersion = "SheetVersion";
			public const string ResourceTable = "ResourceTable";
			public const string TableStartAddress = "TableStartAddress";
		}

		public static class SheetTypes
		{
			public const string ClientRateTables = "Rate Tables";
			public const string GlobalRateTables = "Global Rate Tables";
			public const string GlobalLookupTables = "Global Lookup Tables";
		}
	}

	public static class CalcEngines
	{
		public static class RangeNames
		{
			public const string StartTables = "StartTables";
		}

		public static class SheetTypes
		{
			public const string Input = "Input";
		}

		private static string[] PreviewSheetTypes => new[] { "Result", "ResultXml", "FolderItem" };
		private static string[] ResultSheetTypes => new[] { "Update", "ReportData" }.Concat( PreviewSheetTypes ).ToArray();
		public static bool IsPreviewSheet( string? sheetType ) => !string.IsNullOrEmpty( sheetType ) && PreviewSheetTypes.Contains( sheetType );
		public static bool IsResultSheet( string? sheetType ) => !string.IsNullOrEmpty( sheetType ) && ResultSheetTypes.Contains( sheetType );
		public static bool IsRBLeSheet( string? sheetType ) => !string.IsNullOrEmpty( sheetType ) && new[] { SheetTypes.Input }.Concat( ResultSheetTypes ).Contains( sheetType );
	}
}