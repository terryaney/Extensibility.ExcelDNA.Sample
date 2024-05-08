using System.Text.Json.Nodes;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal class XmlMappingInfo
{
	public required string ClientName { get; init; }
	public required string InputFile { get; init; }
	public required string OutputFile { get; init; }
	public required JsonObject WindowConfiguration { get; init; }
}