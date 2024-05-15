namespace KAT.Camelot.Extensibility.Excel.AddIn.DataExport;

class ExportColumn
{
	public required string Address { get; init; }
	public required ExportField Field { get; init; }

	public bool IsAuthId { get; init; }
	public bool IsDate { get; init; }
	public bool IsDateTime { get; init; }
	public bool IsText { get; init; } // if /text flag, it ignores

	public bool DeleteIfBlank { get; init; }
	public bool AllowReplace { get; init; }
	public string? DateConvertFormat { get; init; }
	public string? Format { get; init; }
	public string? DefaultValue { get; init; }
	public int? DecimalPlacesToInsert { get; init; }
	public bool ToUpper { get; init; }
	public bool ToLower { get; init; }
	public bool IgnoreZero { get; init; }

	public bool IsExportControl { get; init; }
	public bool IsDeleteControl { get; init; }
	public bool IsProfileNote { get; init; }

	public string? RowExportHistoryTable { get; init; }
	public bool IsRowExportHistoryIndex { get; init; }
	public bool IsRowExportHistoryNewIndex { get; init; }
	public bool ClearRowExportHistoryBeforeLoad { get; init; }

	public int Ordinal { get; init; }
	public string? Label { get; init; }
	public bool IgnoreColumn { get; init; }
}