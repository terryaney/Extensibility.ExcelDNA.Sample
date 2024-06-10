using System.Text.Json.Nodes;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public enum ExportFormatType
{
	Csv,
	CsvTransposed,
	Xml
}

internal class LocalBatchInfo
{
	public required string InputFile { get; init; }
	public required string OutputFile { get; init; }
	public required string? Filter { get; init; }
	public required string InputTab { get; init; }
	public required string ResultTab { get; init; }
	public required ExportFormatType ExportType { get; init; }
	public required int? InputRows { get; init; }
	public required int? ErrorCalcEngines { get; init; }
	public required JsonObject WindowConfiguration { get; init; }
}