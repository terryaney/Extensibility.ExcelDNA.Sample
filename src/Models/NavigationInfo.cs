using System.Text.Json.Nodes;

namespace KAT.Extensibility.Excel.AddIn;

internal class NavigationInfo
{
	public required string Target { get; init; }
	public required JsonObject WindowConfiguration { get; init; }
}