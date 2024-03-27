namespace KAT.Extensibility.Excel.AddIn;

public class AddInSettings
{
	public bool ShowRibbon { get; init; }
	public DataService[] DataServices { get; init; } = Array.Empty<DataService>();
	public string? SaveHistoryName { get; init; }
	public CalcEngineManagement CalcEngineManagement { get; init; } = new();
	public DataExport DataExport { get; init; } = new();
	public Features Features { get; init; } = new();
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

public class CalcEngineManagement
{
	public string? Email { get; init; }
	public string? Password { get; init; }
}

public class Features
{
	internal const string Salt = "0fbc569b-f5f9-4a72-8127-ea0a558af5dd";
	public string? SpecSheet { get; init; }
	public string? GlobalTables { get; init; }
}