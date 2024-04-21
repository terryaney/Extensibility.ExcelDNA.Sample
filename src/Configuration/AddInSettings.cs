using System.Net.NetworkInformation;
using KAT.Camelot.Domain.Security.Cryptography;

namespace KAT.Extensibility.Excel.AddIn;

public class AddInSettings
{
	public bool ShowRibbon { get; init; }
	public string ApiEndpoint { get; init; } = null!;
	public string[] DataServices { get; init; } = Array.Empty<string>();
	public string? SaveHistoryName { get; init; }
	public DataExport DataExport { get; init; } = new();
	public Features Features { get; init; } = new();

	public string? KatUserName { get; set; }
	public string? KatPassword { get; set; }

	public void SetCredentials( string? userName, string? password )
	{
		KatUserName = userName;
		KatPassword = password;
		clearPassword = null;
	}

	private string? clearPassword = null;
	public async Task<string?> GetClearPasswordAsync()
	{
		if ( clearPassword != null ) return clearPassword;

		var macAddress = GetMacAddress();

		if ( KatPassword == null || macAddress == null ) return null;

		var decryptedSetting = await Cryptography3DES.DefaultDecryptAsync( KatPassword );
		var macAddressHash = Hash.SHA256Hash( macAddress );
		
		if ( !decryptedSetting.StartsWith( macAddressHash ) ) return null;

		return clearPassword = await Cryptography3DES.DefaultDecryptAsync( decryptedSetting[ macAddressHash.Length.. ] );
	}

	public static async Task<string?> EncryptPasswordAsync( string password )
	{
		var macAddress = GetMacAddress();

		if ( macAddress == null ) return null;

		var encryptedPassword = await Cryptography3DES.DefaultEncryptAsync( password );
		var macAddressHash = Hash.SHA256Hash( macAddress );
		return await Cryptography3DES.DefaultEncryptAsync( macAddressHash + encryptedPassword );
	}

	private static string? GetMacAddress()
	{
		foreach ( var nic in NetworkInterface.GetAllNetworkInterfaces() )
		{
			// Only consider Ethernet network interfaces
			if ( nic.NetworkInterfaceType == NetworkInterfaceType.Ethernet && nic.OperationalStatus == OperationalStatus.Up )
			{
				return nic.GetPhysicalAddress().ToString();
			}
		}
		return null;
	}
}

public class DataExport
{
	public string? Path { get; init; }
	public bool AppendDateToName { get; init; } = false;

}

public class Features
{
	internal const string Salt = "0fbc569b-f5f9-4a72-8127-ea0a558af5dd";
	public string? ShowDeveloperExports { get; init; }
	public string? GlobalTables { get; init; }
	public string? CalcEngineManagement { get; init; }
}