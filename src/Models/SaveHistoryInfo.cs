using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Extensibility.Excel.AddIn;

internal class SaveHistoryInfo
{
	public required DialogResult Result { get; init; }
	public string Version { get; init; } = null!;
	public string? Description { get; init; }
	public string Author { get; init; } = null!;
	public string UserName { get; init; } = null!;
	public string Password { get; init; } = null!;
	public bool ForceUpload { get; init; }

	public MSExcel.Range HistoryRange { get; init; } = null!;
	public MSExcel.Range VersionRange { get; init; } = null!;
}