using System.Net.NetworkInformation;
using System.Security.Cryptography;
using KAT.Camelot.Domain.Security.Cryptography;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

public partial class AddInSettings
{
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
			if ( 
				( nic.NetworkInterfaceType == NetworkInterfaceType.Ethernet || nic.NetworkInterfaceType == NetworkInterfaceType.Wireless80211 ) && 
				nic.OperationalStatus == OperationalStatus.Up 
			)
			{
				return nic.GetPhysicalAddress().ToString();
			}
		}
		throw new InvalidOperationException( "No Ethernet network interfaces found." );
	}
}