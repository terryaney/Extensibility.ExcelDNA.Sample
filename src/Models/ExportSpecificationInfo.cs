using System.Text.Json.Nodes;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal class ExportSpecificationInfo
{
	public required SpecificationLocation[] Locations { get; init; }
	public bool SaveSpecification { get; init; }
	public required JsonObject WindowConfiguration { get; init; }
}

internal class SpecificationLocation
{
	public required string Location { get; init; }
	public bool Selected { get; init; }
}