namespace KAT.Extensibility.Excel.AddIn;

internal partial class Credentials : Form
{
	public Credentials()
	{
		InitializeComponent();
	}

	public CredentialInfo? GetCredentials( string? userName, string? password )
	{
		tUserName.Text = userName;
		tPassword.Text = password;
		
		var dialogResult = ShowDialog();

		if ( dialogResult != DialogResult.OK )
		{
			return null;
		}

		return new()
		{
			UserName = tUserName.Text,
			Password = tPassword.Text
		};
	}

	private void Ok_Click( object sender, EventArgs e )
	{
		errorProvider.Clear();

		var isValid = IsValid();

		if ( isValid )
		{
			DialogResult = DialogResult.OK;
			Close();
		}
	}

	private bool IsValid()
	{
		var isValid = true;

		if ( string.IsNullOrEmpty( tUserName.Text ) )
		{
			errorProvider.SetError( tUserName, "You must provide a User Name to continue." );
			isValid = false;
		}
		if ( string.IsNullOrEmpty( tPassword.Text ) )
		{
			errorProvider.SetError( tPassword, "You must provide a Password to continue." );
			isValid = false;
		}

		return isValid;
	}
}