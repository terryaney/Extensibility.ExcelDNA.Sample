namespace KAT.Extensibility.Excel.AddIn;

internal class CalcEngineUploadInfo
{
	public required string UserName { get; init; }
	public required string Password { get; init; }
	public required string? ExpectedVersion { get; init; }
	public bool ForceUpload { get; init; }
}