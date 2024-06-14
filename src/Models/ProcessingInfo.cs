using System.Text.Json.Nodes;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal class ProcessingInfo
{
	public required DialogResult Result { get; init; }
	public required JsonObject WindowConfiguration { get; init; }
}
