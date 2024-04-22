using System.Text.Json.Nodes;

namespace KAT.Extensibility.Excel.AddIn;

internal partial class Credentials : Form
{
	private readonly JsonObject windowConfiguration;

	public Credentials( JsonObject? windowConfiguration )
	{
		InitializeComponent();
		this.windowConfiguration = windowConfiguration ?? new JsonObject();
	}

	public CredentialInfo? GetCredentials( string? userName, string? password )
	{
		tUserName.Text = userName;
		tPassword.Text = password;

		Location = new Point { X = (int?)windowConfiguration[ "left" ] ?? Left, Y = (int?)windowConfiguration[ "top" ] ?? Top };
		Size = new Size { Width = (int?)windowConfiguration[ "width" ] ?? Width, Height = (int?)windowConfiguration[ "height" ] ?? Height };

		var dialogResult = ShowDialog();

		if ( dialogResult != DialogResult.OK )
		{
			return null;
		}

		windowConfiguration[ "top" ] = Location.Y;
		windowConfiguration[ "left" ] = Location.X;
		windowConfiguration[ "height" ] = Size.Height;
		windowConfiguration[ "width" ] = Size.Width;

		return new()
		{
			UserName = tUserName.Text,
			Password = tPassword.Text,
			WindowConfiguration = windowConfiguration
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