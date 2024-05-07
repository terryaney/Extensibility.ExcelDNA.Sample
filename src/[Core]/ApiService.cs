using System.IO.Compression;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using KAT.Camelot.Abstractions.Api.Contracts.Excel.V1;
using KAT.Camelot.Abstractions.Api.Contracts.Excel.V1.Requests;
using KAT.Camelot.Abstractions.Api.Contracts.Excel.V1.Responses;
using KAT.Camelot.Domain.Extensions;
using KAT.Camelot.Domain.Services;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public class ApiService
{
	private readonly IHttpClientFactory httpClientFactory;
	private readonly IxDSRepository xDSRepository;

	public ApiService( IHttpClientFactory httpClientFactory, IxDSRepository xDSRepository )
	{
		this.httpClientFactory = httpClientFactory;
		this.xDSRepository = xDSRepository;
	}

	public async Task<ApiResponse<string>> GetSpreadsheetGearLicenseAsync( string? userName, string? password )
	{
		var url = $"{AddIn.Settings.ApiEndpoint}{ApiEndpoints.Utility.SpreadsheetGear}";
		
		var (httpResponse, validations) = await SendHttpRequestAsync( userName, password, url, HttpMethod.Get );
		
		if ( validations != null )
		{
			return new() { Validations = validations };
		}

		using var response = httpResponse!;

		return new() { Response = await response.Content.ReadAsStringAsync() };
	}

	public async Task<ApiResponse<IEnumerable<DebugFile>>> GetDebugFilesAsync( string calcEngine, string? userName, string? password )
	{
		var url = $"{AddIn.Settings.ApiEndpoint}{ApiEndpoints.CalcEngines.Build.DebugListing( calcEngine )}";
		var (httpResponse, validations) = await SendRequestAsync<DebugFile[]>( userName, password, url, HttpMethod.Get );
		return new() { Response = httpResponse, Validations = validations };
	}

	public async Task<ApiResponse<CalcEngineInfo>> GetCalcEngineInfoAsync( string calcEngine, string? userName, string? password )
	{
		var url = $"{AddIn.Settings.ApiEndpoint}{ ApiEndpoints.CalcEngines.Build.Get( Path.GetFileNameWithoutExtension( calcEngine ) )}";
		var (httpResponse, validations) = await SendRequestAsync<CalcEngineInfo>( userName, password, url, HttpMethod.Get );
		return new() { Response = httpResponse, Validations = validations };
	}

	public async Task<ApiResponse<string>> DownloadDebugAsync( string calcEngine, int versionKey, string? userName, string? password )
	{
		var url = $"{AddIn.Settings.ApiEndpoint}{ ApiEndpoints.CalcEngines.Build.DebugDownload( Path.GetFileNameWithoutExtension( calcEngine ), versionKey )}";
		
		var (httpResponse, validations) = await SendHttpRequestAsync( userName, password, url, HttpMethod.Get );
		
		if ( validations != null )
		{
			return new() { Validations = validations };
		}
		
		using var response = httpResponse!;

		var downloadFolder = Path.Combine( Environment.GetFolderPath( Environment.SpecialFolder.UserProfile ), "Downloads" );
		var fileName = Path.Combine( downloadFolder, response.Content.Headers.ContentDisposition!.FileName!.Replace( "\"", "" ) );

		using var source = await response.Content.ReadAsStreamAsync();
		using var dest = File.Create( fileName );
		await source.CopyToAsync( dest );

		return new() { Response = fileName };
	}

	public async Task<ApiValidation[]?> DownloadLatestAsync( string path, string? userName, string? password )
	{
		var calcEngine = Path.GetFileNameWithoutExtension( path );
		var url = $"{AddIn.Settings.ApiEndpoint}{ ApiEndpoints.CalcEngines.Build.DownloadLatest( calcEngine )}";
		
		var (httpResponse, validations) = await SendHttpRequestAsync( userName, password, url, HttpMethod.Get );

		if ( validations != null )
		{
			return validations;
		}

		using var response = httpResponse!;

		Directory.CreateDirectory( Path.GetDirectoryName( path )! );

		using var source = await response.Content.ReadAsStreamAsync();
		using var dest = File.Create( path );
		await source.CopyToAsync( dest );

		return null;
	}

	public async Task CheckinAsync( string calcEngine, string? userName, string? password )
	{
		var url = $"{AddIn.Settings.ApiEndpoint}{ ApiEndpoints.CalcEngines.Build.Checkin( Path.GetFileNameWithoutExtension( calcEngine ) )}";
		await SendRequestWithoutResponseAsync( userName, password, url, HttpMethod.Patch );
	}

	public async Task CheckoutAsync( string calcEngine, string? userName, string? password )
	{
		var url = $"{AddIn.Settings.ApiEndpoint}{ ApiEndpoints.CalcEngines.Build.Checkout( Path.GetFileNameWithoutExtension( calcEngine ) )}";
		await SendRequestWithoutResponseAsync( userName, password, url, HttpMethod.Patch );
	}

	public async Task<ApiValidation[]?> WaitForEmailBlastAsync( string token, string userName, string password )
	{
		var url = $"{AddIn.Settings.ApiEndpoint}{ ApiEndpoints.Utility.Build.WaitEmailBlastComplete( token ) }";
		return await SendRequestWithoutResponseAsync( userName, password, url, HttpMethod.Get );
	}

	public async Task<ApiResponse<string>> EmailBlastAsync( string contentFile, EmailBlastRequest reqeust, string? userName, string? password, CancellationToken cancellationToken = default )
	{
		var ms = new MemoryStream();
		using ( var zip = new ZipArchive( ms, ZipArchiveMode.Create, true ) )
		{
			var emailBody = await File.ReadAllTextAsync( contentFile, cancellationToken: cancellationToken );

			var currentDirectory = Path.GetDirectoryName( contentFile )!;
			var matches = Regex.Matches(emailBody, @"<img[^>]+src=(?<src>[""'][^""']+[""'])[^>]*>");
			var imageAttachments = new List<EmailBlastAttachment>();

			foreach ( var match in matches.Cast<Match>() )
			{
				var originalSrc = match.Groups[ "src" ].Value;
				var src = originalSrc.Trim('"', '\'');

				if ( !src.StartsWith( "http", StringComparison.OrdinalIgnoreCase ) )
				{
					var contentId = Guid.NewGuid().ToString();
					var fi = new FileInfo( Path.Combine( currentDirectory, src ) );

					var attachment = fi.Exists
						? new EmailBlastAttachment { Id = contentId, Name = fi.FullName, ContentId = contentId }
						: new EmailBlastAttachment { Id = contentId, Name = src };

					src = fi.Exists ? $"cid:{contentId}" : src;

					imageAttachments.Add( attachment );
				}

				emailBody = emailBody.Replace( originalSrc, $"\"{src}\"" );
			}

			if ( imageAttachments.Any( a => a.ContentId == null ) )
			{
				return new() 
				{ 
					Validations = 
						imageAttachments.Where( a => a.ContentId == null )
							.Select( a => new ApiValidation { Name = "Image Attachment", Message = $"{a.Name} is used as an image source but could not be found." } )
							.ToArray()
				};
			}

			reqeust.Attachments = reqeust.Attachments.Concat( imageAttachments ).ToArray();

			var entry = zip.CreateEntry( ".content.html", CompressionLevel.Fastest );
			using var contentStream = entry.Open();
			using var contentMs = new MemoryStream( Encoding.UTF8.GetBytes( emailBody ) );
			await contentMs.CopyToAsync( contentStream, cancellationToken );
			await contentStream.DisposeAsync();
			await contentMs.DisposeAsync();

			entry = zip.CreateEntry( ".configuration.json", CompressionLevel.Fastest );
			using var es = entry.Open();
			JsonSerializer.Serialize( es, reqeust, new JsonSerializerOptions { WriteIndented = true } );
			await es.DisposeAsync();

			foreach ( var attachment in reqeust.Attachments )
			{
				entry = zip.CreateEntry( attachment.Id, CompressionLevel.Fastest );
				using var attachmentStream = entry.Open();
				using var afs = File.OpenRead( attachment.Name );
				await afs.CopyToAsync( attachmentStream, cancellationToken );
				await attachmentStream.DisposeAsync();
				await afs.DisposeAsync();
			}
		}

		ms.Position = 0;

		using var form = new MultipartFormDataContent
		{
			{ new StreamContent( ms ), "file", "emailBlast.zip" }
		};

		var url = $"{AddIn.Settings.ApiEndpoint}{ ApiEndpoints.Utility.EmailBlast }";

		var (httpResponse, validations) = await SendHttpRequestAsync( userName, password, url, HttpMethod.Post, form );
		
		if ( validations != null )
		{
			return new() { Validations = validations };
		}

		using var response = httpResponse!;

		var token = await response.Content.ReadAsStringAsync( cancellationToken );

		return new () { Response = token };
	}

	public async Task<ApiValidation[]?> UpdateGlobalTablesAsync( string? clientName, string[] targets, JsonObject globalTables, string? userName, string? password, CancellationToken cancellationToken = default )
	{
		var client = clientName ?? "Global";
		if ( targets.Any( t => string.Compare( t, "LOCAL", true ) == 0 ) )
		{
			if ( !( userName?.Contains( '@' ) ?? false ) )
			{
				return CredentialsMissing;
			}
			await xDSRepository.UpdateGlobalLookupsAsync( client, globalTables, userName[ userName.IndexOf( "@" ).. ], cancellationToken );
		}

		var remotes = targets.Where( t => string.Compare( t, "LOCAL", true ) != 0 );

		if ( remotes.Any() )
		{
			var ms = new MemoryStream();
			using ( var zip = new ZipArchive( ms, ZipArchiveMode.Create, true ) )
			{
				var entry = zip.CreateEntry( "globalTables.json", CompressionLevel.Fastest );

				using var es = entry.Open();
				JsonSerializer.Serialize( es, globalTables, new JsonSerializerOptions { WriteIndented = true } );
			}

			ms.Position = 0;

			using var form = new MultipartFormDataContent
			{
				{ new StreamContent( ms ), "file", "globalTables.zip" },
				{ new StringContent( client ), "clientName" },
				{ new StringContent( $"[ {string.Join( ", ", remotes.Select( t => $"\"{t}\"" ))} ]" ), "targets" }
			};

			var url = $"{AddIn.Settings.ApiEndpoint}{ ApiEndpoints.xDSData.GlobalTables }";

			return await SendRequestWithoutResponseAsync( userName, password, url, HttpMethod.Post, form );
		}

		return null;
	}

	private async Task<(T? Response, ApiValidation[]? Validations)> SendRequestAsync<T>( string? userName, string? password, string url, HttpMethod method ) where T : class
	{
		var (httpResponse, validations) = await SendHttpRequestAsync( userName, password, url, method );

		if ( validations != null )
		{
			return (null, validations);
		}

		using var response = httpResponse!;
		try
		{
			return (await response.Content.ReadFromJsonAsync<T>(), null);
		}
		catch ( Exception ex )
		{
			throw new ApplicationException( $"Unable to parse response from {url}.", ex );
		}
	}

	private static ApiValidation[] CredentialsMissing => new[] { new ApiValidation { Name = "userName", Message = "You must provide your KAT Management credentials to use this functionality." } };

	private async Task<ApiValidation[]?> SendRequestWithoutResponseAsync( string? userName, string? password, string url, HttpMethod method, HttpContent? content = null ) =>
		( await SendHttpRequestAsync( userName, password, url, method, content ) ).Validations;

	private async Task<(HttpResponseMessage? Response, ApiValidation[]? Validations)> SendHttpRequestAsync( string? userName, string? password, string url, HttpMethod method, HttpContent? content = null )
	{
		if ( string.IsNullOrEmpty( userName ) || string.IsNullOrEmpty( password ) )
		{
			return (null, CredentialsMissing);
		}

		using var httpClient = httpClientFactory.CreateClient();
		
		httpClient.DefaultRequestHeaders.Add( "x-kat-email", userName );
		httpClient.DefaultRequestHeaders.Add( "x-kat-password", password );

		using var request = new HttpRequestMessage( method, url ) { Content = content };

		HttpResponseMessage? response = null;
		try
		{
			response = await httpClient.SendConduentAsync( request );

			if ( response.StatusCode == System.Net.HttpStatusCode.BadRequest )
			{
				var problemDetails = ( await response.Content.ReadFromJsonAsync<JsonNode>() )!;

				if ( problemDetails[ "errors" ] is JsonObject errors )
				{
					var validations =
						errors.SelectMany( e => 
							( e.Value as JsonArray ?? new JsonArray { e.Value } )
								.Select( m => new ApiValidation { Name = e.Key, Message = (string)m! } )
						).ToArray();
					return (null, validations);
				}

				return (null, new[] { new ApiValidation { Name = "Excel Api", Message = (string)problemDetails[ "detail" ]! } });
			}

			response.EnsureSuccessStatusCode();

			return (response, null);
		}
		catch ( Exception ex )
		{
			string? result = null;
			if ( response != null )
			{
				result = await response.Content.ReadAsStringAsync();
				Console.WriteLine( result );
			}

			throw new ApplicationException( $"Unable to send request to {url}.  Response: {result ?? "Not Available."}", ex );
		}
	}
}