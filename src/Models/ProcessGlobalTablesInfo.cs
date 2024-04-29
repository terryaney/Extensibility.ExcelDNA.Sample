using System.Text.Json.Nodes;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal class ProcessGlobalTablesInfo
{
	public required string[] Targets { get; init; }
	public required string UserName { get; init; }
	public required string Password { get; init; }
	public required JsonObject WindowConfiguration { get; init; }
}