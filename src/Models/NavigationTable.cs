namespace KAT.Camelot.Extensibility.Excel.AddIn;

public class NavigationTable : NavigationTarget
{
	public string? Description { get; init; }
	public required NavigationTarget[] Columns { get; init; }
}

public class NavigationTarget
{
	public required string Name { get; init; }
	public required string Address { get; init; }
}