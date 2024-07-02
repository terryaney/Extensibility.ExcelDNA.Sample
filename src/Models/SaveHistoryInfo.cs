using System.Text.Json.Nodes;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal class SaveHistoryInfo
{
	public required DialogResult Result { get; init; }
	public string Version { get; init; } = null!;
	public string? Description { get; init; }
	public string Author { get; init; } = null!;
	public string UserName { get; init; } = null!;
	public string Password { get; init; } = null!;

	public MSExcel.Range HistoryRange { get; init; } = null!;
	public MSExcel.Range VersionRange { get; init; } = null!;
	public required JsonObject WindowConfiguration { get; init; }
}