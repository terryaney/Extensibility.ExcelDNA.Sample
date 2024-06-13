using System.Net.NetworkInformation;
using System.Security.Cryptography;
using KAT.Camelot.Domain.Security.Cryptography;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public class AddInSettings
{
	public bool ShowRibbon { get; init; }
	// Path (or file if path is in %PATH%) to the text editor to use for opening xml/json files...
	public string TextEditor { get; init; } = @"C:\Program Files\Microsoft VS Code\bin\code.exe";
	public string ApiEndpoint { get; init; } = null!;
	public string[] DataServices { get; init; } = Array.Empty<string>();
	public string[] SpecificationFileLocations { get; init; } = Array.Empty<string>();
	
	public string? SaveHistoryName { get; init; }
	public DataExportSettings DataExport { get; init; } = new();
	public Help Help { get; init; } = new();

	public string? KatUserName { get; set; }
	public string? KatPassword { get; set; }



	public async Task<string?> SetCredentialsAsync( string userName, string password )
	{
		KatUserName = userName;

		var macAddress = GetMacAddress();
		var encryptedPassword = await Cryptography3DES.DefaultEncryptAsync( password );
		var macAddressHash = Hash.SHA256Hash( macAddress );
		return KatPassword = await Cryptography3DES.DefaultEncryptAsync( macAddressHash + encryptedPassword );
	}

	string? clearPassword;
	public async Task<string?> GetClearPasswordAsync()
	{
		if ( clearPassword != null ) return clearPassword;
		if ( KatPassword == null ) return null;
		if ( KatPassword.Length < 45 )
		{
			clearPassword = KatPassword; // password is in clear text
			return clearPassword;
		}
		
		try
		{
			var macAddress = GetMacAddress();
			var decryptedSetting = await Cryptography3DES.DefaultDecryptAsync( KatPassword );
			var macAddressHash = Hash.SHA256Hash( macAddress );
			
			if ( !decryptedSetting.StartsWith( macAddressHash ) ) return null;

			return clearPassword = await Cryptography3DES.DefaultDecryptAsync( decryptedSetting[ macAddressHash.Length.. ] );
		}
		catch ( CryptographicException ex ) when ( ex.Message == "The input data is not a complete block." )
		{
			return KatPassword; // password is in clear text (debugging)
		}
	}

	private static string GetMacAddress()
	{
		foreach ( var nic in NetworkInterface.GetAllNetworkInterfaces() )
		{
			// Only consider Ethernet network interfaces
			if ( nic.NetworkInterfaceType == NetworkInterfaceType.Ethernet && nic.OperationalStatus == OperationalStatus.Up )
			{
				return nic.GetPhysicalAddress().ToString();
			}
		}
		throw new InvalidOperationException( "No Ethernet network interfaces found." );
	}
}

public class DataExportSettings
{
	public string? Path { get; init; }
	public bool AppendDateToName { get; init; } = false;

}

public class Help
{
	public string Url { get; init; } = "https://github.com/terryaney/Extensibility.ExcelDNA.Sample";
	public string? OfflineUrl { get; init; }
	public bool Offline { get; init; }
	public string GetOfflineUrl() => OfflineUrl ?? "file:///" + Path.Combine( AddIn.XllPath, "Resources", "Help", "readme.md" );
}