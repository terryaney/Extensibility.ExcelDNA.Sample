using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using KAT.Camelot.Abstractions.Api.Contracts.Excel.V1;
using KAT.Camelot.Abstractions.Api.Contracts.Excel.V1.Requests;
using KAT.Camelot.Abstractions.Api.Contracts.Excel.V1.Responses;
using KAT.Camelot.Domain.Extensions;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public class ApiService
{
	private readonly IHttpClientFactory httpClientFactory;

	public ApiService( IHttpClientFactory httpClientFactory )
	{
		this.httpClientFactory = httpClientFactory;
	}

	public async Task<IEnumerable<DebugFile>> GetDebugFilesAsync( string calcEngine, string? userName, string? password )
	{
		var url = $"{AddIn.Settings.ApiEndpoint}{ApiEndpoints.CalcEngines.Build.DebugListing( calcEngine )}";
		return await SendRequestAsync<DebugFile[]>( userName, password, url, HttpMethod.Get ) ?? Array.Empty<DebugFile>();
	}

	public async Task<CalcEngineInfo?> GetCalcEngineInfoAsync( string calcEngine, string? userName, string? password )
	{
		var url = $"{AddIn.Settings.ApiEndpoint}{ ApiEndpoints.CalcEngines.Build.Get( Path.GetFileNameWithoutExtension( calcEngine ) )}";
		return await SendRequestAsync<CalcEngineInfo>( userName, password, url, HttpMethod.Get );
	}

	public async Task<string?> DownloadDebugAsync( string calcEngine, int versionKey, string? userName, string? password )
	{
		if ( string.IsNullOrEmpty( userName ) || string.IsNullOrEmpty( password ) )
		{
			return null;
		}

		var url = $"{AddIn.Settings.ApiEndpoint}{ ApiEndpoints.CalcEngines.Build.DebugDownload( Path.GetFileNameWithoutExtension( calcEngine ), versionKey )}";
		using var response = await SendHttpRequestAsync( userName, password, url, HttpMethod.Get );

		var downloadFolder = Path.Combine( Environment.GetFolderPath( Environment.SpecialFolder.UserProfile ), "Downloads" );
		var fileName = Path.Combine( downloadFolder, response.Content.Headers.ContentDisposition!.FileName!.Replace( "\"", "" ) );

		using var source = await response.Content.ReadAsStreamAsync();
		using var dest = File.Create( fileName );
		await source.CopyToAsync( dest );

		return fileName;
	}

	public async Task<bool> DownloadLatestAsync( string path, string? userName, string? password )
	{
		if ( string.IsNullOrEmpty( userName ) || string.IsNullOrEmpty( password ) )
		{
			return false;
		}

		var calcEngine = Path.GetFileNameWithoutExtension( path );
		var url = $"{AddIn.Settings.ApiEndpoint}{ ApiEndpoints.CalcEngines.Build.DownloadLatest( calcEngine )}";
		using var response = await SendHttpRequestAsync( userName, password, url, HttpMethod.Get );

		Directory.CreateDirectory( Path.GetDirectoryName( path )! );

		using var source = await response.Content.ReadAsStreamAsync();
		using var dest = File.Create( path );
		await source.CopyToAsync( dest );

		return true;
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

	private async Task<T?> SendRequestAsync<T>( string? userName, string? password, string url, HttpMethod method ) where T : class
	{
		if ( string.IsNullOrEmpty( userName ) || string.IsNullOrEmpty( password ) )
		{
			return null;
		}

		using var response = await SendHttpRequestAsync( userName, password, url, method );
		try
		{
			return await response.Content.ReadFromJsonAsync<T>();
		}
		catch ( Exception ex )
		{
			throw new ApplicationException( $"Unable to parse response from {url}.", ex );
		}
	}

	private async Task SendRequestWithoutResponseAsync( string? userName, string? password, string url, HttpMethod method )
	{
		if ( string.IsNullOrEmpty( userName ) || string.IsNullOrEmpty( password ) )
		{
			return;
		}
		await SendHttpRequestAsync( userName, password, url, method );
	}

	private async Task<HttpResponseMessage> SendHttpRequestAsync( string userName, string password, string url, HttpMethod method )
	{
		using var httpClient = httpClientFactory.CreateClient();
		
		httpClient.DefaultRequestHeaders.Add( "x-kat-email", userName );
		httpClient.DefaultRequestHeaders.Add( "x-kat-password", password );
		
		using var request = new HttpRequestMessage( method, url );

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