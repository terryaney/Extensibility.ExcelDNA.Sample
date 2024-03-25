namespace KAT.Extensibility.Excel.AddIn;

public class AddInSettings
{
	public bool ShowRibbon { get; init; }
	public DataService[] DataServices { get; init; } = Array.Empty<DataService>();
	public SaveHistory SaveHistory { get; init; } = new();
	public DataExport DataExport { get; init; } = new();
}

public class DataService
{
	public required string Name { get; init; }
	public required string Url { get; init; }
}

public class DataExport
{
	public string? Path { get; init; }
	public bool AppendDateToName { get; init; } = false;

}

public class SaveHistory
{
	public bool Enabled { get; init; } = true;
	public string? Name { get; init; }
	public string? Email { get; init; }
	public string? Password { get; init; }

}