using System.Text.Json.Nodes;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

enum ExportDataAction
{
	Export,
	Validate
}

internal class ExportDataInfo
{
	public required ExportDataAction Action { get; init; }
	public string? ClientName { get; init; }
	public string? AuthIdToExport { get; init; }
	public string? OutputFile { get; init; }
	public int? MaxFileSize { get; init; }
	public required JsonObject WindowConfiguration { get; init; }
}