using System.Text.Json.Nodes;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal partial class ProcessGlobalTables : Form
{
	private readonly bool requireClient;
	private readonly JsonObject windowConfiguration;

	public ProcessGlobalTables( bool requireClient, string[] targets, JsonObject? windowConfiguration )
	{
		InitializeComponent();

		this.targets.Items.Clear();
		this.targets.Items.AddRange( targets );
		if ( targets.Contains( "LOCAL" ) )
		{
			this.targets.SetItemChecked( this.targets.Items.IndexOf( "LOCAL" ), true );
		}
		this.requireClient = requireClient;
		this.windowConfiguration = windowConfiguration ?? new JsonObject();
	}

	public ProcessGlobalTablesInfo? GetInfo( string? userName, string? password )
	{
		targets.Select();
		emailAddress.Text = userName;
		this.password.Text = password;
		clientName.Enabled = clientNameLabel.Enabled = requireClient;

		var dialogResult = ShowDialog();

		if ( dialogResult != DialogResult.OK )
		{
			return null;
		}

		windowConfiguration[ "state" ] = WindowState.ToString();

		if ( WindowState == FormWindowState.Normal )
		{
			windowConfiguration[ "top" ] = Location.Y;
			windowConfiguration[ "left" ] = Location.X;
			windowConfiguration[ "height" ] = Size.Height;
			windowConfiguration[ "width" ] = Size.Width;
		}

		return new()
		{
			ClientName = requireClient ? clientName.Text : null,
			Targets = targets.CheckedItems.Cast<string>().ToArray(),
			UserName = emailAddress.Text,
			Password = this.password.Text,
			WindowConfiguration = windowConfiguration
		};
	}

	private void ProcessGlobalTable_Load( object sender, EventArgs e )
	{
		WindowState = Enum.TryParse( (string?)windowConfiguration[ "state" ], out FormWindowState state) ? state : FormWindowState.Normal;
		Location = new Point { X = (int?)windowConfiguration[ "left" ] ?? Left, Y = (int?)windowConfiguration[ "top" ] ?? Top };
		Size = new Size { Width = (int?)windowConfiguration[ "width" ] ?? Width, Height = (int?)windowConfiguration[ "height" ] ?? Height };
	}

	private void Ok_Click( object sender, EventArgs e )
	{
		errorProvider.Clear();

		if ( requireClient && string.IsNullOrEmpty( clientName.Text ) )
		{
			errorProvider.SetError( clientName, "You must provide an Client Name to continue." );
		}
		if ( string.IsNullOrEmpty( emailAddress.Text ) )
		{
			errorProvider.SetError( emailAddress, "You must provide an Email Address to continue." );
		}
		if ( string.IsNullOrEmpty( password.Text ) )
		{
			errorProvider.SetError( password, "You must provide a Password to continue." );
		}
		if ( targets.CheckedItems.Count == 0 )
		{
			errorProvider.SetError( targets, "You must select at least one Target to continue." );
		}

		if ( !errorProvider.HasErrors )
		{
			DialogResult = DialogResult.OK;
			Close();
		}
	}
}