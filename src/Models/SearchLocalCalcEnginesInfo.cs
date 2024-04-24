using System.Text.Json.Nodes;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal class SearchLocalCalcEnginesInfo
{
	public required string Folder { get; init; }
	public required string[] Tokens { get; init; }
	public required JsonObject WindowConfiguration { get; init; }
}