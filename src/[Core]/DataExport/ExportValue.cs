namespace KAT.Camelot.Extensibility.Excel.AddIn.DataExport;

class ExportValue : ExportField
{
	public bool IsAuthID { get; set; }
	public string? Value { get; set; }
	public bool Clear { get; set; }
	public bool SkipExport { get; set; }
	public bool AllowReplace { get; set; }

	public string? Subject { get; set; }
	public string? Body { get; set; }
}