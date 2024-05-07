using System.Text.Json.Nodes;
using KAT.Camelot.Abstractions.Api.Contracts.Excel.V1.Requests;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal class EmailBlastInfo : EmailBlastRequest
{
	public required JsonObject WindowConfiguration { get; init; }
	public required string Body { get; init; }
	public required string UserName { get; init; }
	public required string Password { get; init; }
}

internal class EmailBlastRequestInfo
{
	public required int AddressesPerEmail { get; init; }
	public required int WaitPerBatch { get; init; }
	public required string? Bcc { get; init; }
	public required string? From { get; init; }
	public required string? Subject { get; init; }
	public required string? Body { get; init; }
	public required string[] Attachments { get; init; }
}