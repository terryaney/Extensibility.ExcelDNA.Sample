using System.Net.NetworkInformation;

namespace KAT.Extensibility.Excel.AddIn;

public class AddInSettings
{
	public bool ShowRibbon { get; init; }
	public string ApiEndpoint { get; init; } = null!;
	public string[] DataServices { get; init; } = Array.Empty<string>();
	public string? SaveHistoryName { get; init; }
	public CalcEngineManagement CalcEngineManagement { get; init; } = new();
	public DataExport DataExport { get; init; } = new();
	public Features Features { get; init; } = new();
}

public class DataExport
{
	public string? Path { get; init; }
	public bool AppendDateToName { get; init; } = false;

}

public class CalcEngineManagement
{
	public string? Email { get; init; }
	public string? Password { get; init; }

	private string? clearPassword = null;
	public async Task<string?> GetClearPasswordAsync()
	{
		if ( clearPassword != null ) return clearPassword;

		var macAddress = GetMacAddress();

		if ( Password == null || macAddress == null ) return null;

		// TODO: var decryptedSetting = await Cryptography3DES.DefaultDecryptAsync( Password );
		// TODO: var macAddressHash = Password.SHA256Hash( macAddress );
		await Task.Delay( 0 );
		var decryptedSetting = Password;
		var macAddressHash = macAddress;
		
		if ( !decryptedSetting.StartsWith( macAddressHash ) ) return null;

		// TODO: return clearPassword = await Cryptography3DES.DefaultDecryptAsync( decryptedSetting[ macAddressHash.Length.. ] );
		return clearPassword = decryptedSetting[ macAddressHash.Length.. ]; 
	}

	public static async Task<string?> EncryptPasswordAsync( string password )
	{
		var macAddress = GetMacAddress();

		if ( macAddress == null ) return null;

		// TODO: var encryptedPassword = await Cryptography3DES.DefaultEncryptAsync( password );
		// TODO: var macAddressHash = Camelot.Domain.Security.Cryptography.Password.SHA256Hash( macAddress );
		// TODO: return await Cryptography3DES.DefaultEncryptAsync( macAddressHash + encryptedPassword );
		await Task.Delay( 0 );
		var encryptedPassword = password;
		var macAddressHash = macAddress;
		return macAddressHash + encryptedPassword;
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

public class Features
{
	internal const string Salt = "0fbc569b-f5f9-4a72-8127-ea0a558af5dd";
	public string? ShowDeveloperExports { get; init; }
	public string? GlobalTables { get; init; }
	public string? CalcEngineManagement { get; init; }
}