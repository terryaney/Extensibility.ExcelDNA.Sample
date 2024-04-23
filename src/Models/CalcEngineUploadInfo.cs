namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal class CalcEngineUploadInfo : CredentialInfo
{
	public required string? ExpectedVersion { get; init; }
	public bool ForceUpload { get; init; }
}