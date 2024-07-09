namespace KAT.Camelot.Extensibility.Excel.AddIn;

public partial class AddInSettings
{
	public bool ShowRibbon { get; init; }
	public bool ShowConfigurationExports { get; init; }
	// Path (or file if path is in %PATH%) to the text editor to use for opening xml/json files...
	public string TextEditor { get; init; } = @"C:\Program Files\Microsoft VS Code\code.exe";
	public string ApiEndpoint { get; init; } = "https://btr.lifeatworkportal.com/services/camelot/excel";
	public string[] DataServices { get; init; } = Array.Empty<string>();
	public string[] SpecificationFileLocations { get; init; } = Array.Empty<string>();
	
	public string? SaveHistoryName { get; init; }
	public DataExportSettings DataExport { get; init; } = new();
	public Help Help { get; init; } = new();

	public string? KatUserName { get; set; }
	public string? KatPassword { get; set; }
}

public class DataExportSettings
{
	public string? Path { get; init; }
	public bool AppendDateToName { get; init; } = false;

}

public class Help
{
	public string Url { get; init; } = "https://github.com/terryaney/Documentation.Camelot/blob/main/RBLe/ExcelAddIn.md";
	public string? OfflineUrl { get; init; }
	public bool Offline { get; init; }
	public string GetOfflineUrl() => OfflineUrl ?? "file:///" + Path.Combine( AddIn.XllPath, "Resources", "Help", "readme.md" );
}