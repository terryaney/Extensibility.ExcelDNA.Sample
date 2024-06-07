using System.Text.Json.Nodes;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal class LoadInputTabInfo
{
	public required string DataSource { get; init; }
	public required string ClientName { get; init; }
	public required string? AuthId { get; init; }
	public required bool LoadLookupTables { get; init; }
	public required bool DownloadGlobalTables { get; init; }
	public required string? DataSourceFile { get; init; }
	public required string? ConfigLookupsUrl { get; init; }
	public required string? ConfigLookupsPath { get; init; }
	public required string? UserName { get; init; }
	public required string? Password { get; init; }
	public required JsonObject WindowConfiguration { get; init; }
}