using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using KAT.Camelot.Abstractions.Api.Contracts.Excel.V1;
using KAT.Camelot.Abstractions.Api.Contracts.Excel.V1.Requests;
using KAT.Camelot.Abstractions.Api.Contracts.Excel.V1.Responses;
using KAT.Camelot.Domain.Extensions;

namespace KAT.Extensibility.Excel.AddIn;

// TODO: Get ribbon xml ids standardized and in here?
public class ApiService
{
	private readonly IHttpClientFactory httpClientFactory;

	public ApiService( IHttpClientFactory httpClientFactory )
	{
		this.httpClientFactory = httpClientFactory;
	}

	public async Task<IEnumerable<DebugFile>> GetDebugFilesAsync( string calcEngine, string? userName, string? password )
	{
		var url = $"{AddIn.Settings.ApiEndpoint}{ApiEndpoints.CalcEngines.DebugListing}";
		return await SendRequestAsync<DebugFile[]>( calcEngine, userName, password, url ) ?? Array.Empty<DebugFile>();
	}

	public async Task<CalcEngineInfo?> GetCalcEngineInfoAsync( string calcEngine, string? userName, string? password )
	{
		var url = $"{AddIn.Settings.ApiEndpoint}{ ApiEndpoints.CalcEngines.Build.Get( Path.GetFileNameWithoutExtension( calcEngine ) )}";
		return await SendRequestAsync<CalcEngineInfo>( calcEngine, userName, password, url );
	}

	public async Task<string?> DownloadDebugAsync( int versionKey, string? userName, string? password )
	{
		if ( string.IsNullOrEmpty( userName ) || string.IsNullOrEmpty( password ) )
		{
			return null;
		}

		var url = $"{AddIn.Settings.ApiEndpoint}{ ApiEndpoints.CalcEngines.Download}";
		using var response = await SendHttpRequestAsync(
			new DownloadDebugRequest { VersionKey = versionKey, Email = userName, Password = password },
			url
		);

		var downloadFolder = Path.Combine( Environment.GetFolderPath( Environment.SpecialFolder.UserProfile ), "Downloads" );
		var fileName = Path.Combine( downloadFolder, response.Content.Headers.ContentDisposition!.FileName!.Replace( "\"", "" ) );

		using var source = await response.Content.ReadAsStreamAsync();
		using var dest = File.Create( fileName );
		await source.CopyToAsync( dest );

		return fileName;
	}

	public async Task Checkin( string calcEngine, string? userName, string? password )
	{
		var url = $"{AddIn.Settings.ApiEndpoint}{ ApiEndpoints.CalcEngines.Build.Checkin( Path.GetFileNameWithoutExtension( calcEngine ) )}";
		await SendRequestAsync( calcEngine, userName, password, url );
	}

	public async Task Checkout( string calcEngine, string? userName, string? password )
	{
		var url = $"{AddIn.Settings.ApiEndpoint}{ ApiEndpoints.CalcEngines.Build.Checkout( Path.GetFileNameWithoutExtension( calcEngine ) )}";
		await SendRequestAsync( calcEngine, userName, password, url );
	}

	private async Task<T?> SendRequestAsync<T>( string calcEngine, string? userName, string? password, string url ) where T : class
	{
		if ( string.IsNullOrEmpty( userName ) || string.IsNullOrEmpty( password ) )
		{
			return null;
		}

		using var response = await SendHttpRequestAsync( calcEngine, userName, password, url );
		try
		{
			return await response.Content.ReadFromJsonAsync<T>();
		}
		catch ( Exception ex )
		{
			throw new ApplicationException( $"Unable to parse response from {url}.", ex );
		}
	}

	private async Task SendRequestAsync( string calcEngine, string? userName, string? password, string url )
	{
		if ( string.IsNullOrEmpty( userName ) || string.IsNullOrEmpty( password ) )
		{
			return;
		}
		await SendHttpRequestAsync( calcEngine, userName, password, url );
	}

	private async Task<HttpResponseMessage> SendHttpRequestAsync( string calcEngine, string userName, string password, string url ) =>
		await SendHttpRequestAsync(
			new CalcEngineRequest { Name = calcEngine, Email = userName, Password = password },
			url
		);

	private async Task<HttpResponseMessage> SendHttpRequestAsync<T>( T payload, string url ) where T : class
	{
		using var httpClient = httpClientFactory.CreateClient();
		using var request = new HttpRequestMessage( HttpMethod.Post, url )
		{
			Content = new StringContent( 
				JsonSerializer.Serialize( payload ),
				Encoding.UTF8,
				"application/json"
			)
		};

		// TODO: Global error handling...if ensure success status code throws error, excel crashes...get better global handling...
		try
		{
			var response = await httpClient.SendConduentAsync( request );

			response.EnsureSuccessStatusCode();

			return response;
		}
		catch ( Exception ex )
		{
			throw new ApplicationException( $"Unable to send request to {url}.", ex );
		}
	}
}