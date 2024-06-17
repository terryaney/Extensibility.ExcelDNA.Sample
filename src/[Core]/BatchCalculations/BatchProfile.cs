using KAT.Camelot.Abstractions.RBLe;

namespace KAT.Camelot.Extensibility.Excel.AddIn.BatchCalculations;

class BatchProfile
{
	public required string AuthId { get; init; }
	public required RBLePayload Payload { get; init; }
	public long StreamPosition { get; init; }
}