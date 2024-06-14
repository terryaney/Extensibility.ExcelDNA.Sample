using System.Text.Json.Nodes;
using System.Windows.Forms;
using KAT.Camelot.Domain.Extensions;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal partial class LoadInputTab : Form
{
	public const string LocalProfileFile = "LocalProfileFile";
	public const string LocalGroupFile = "LocalGroupFile";
	private readonly string calcEngine;
	private readonly JsonObject windowConfiguration;
	private bool hasChangedInputFile;

	public LoadInputTab( string? clientName, string? authId, string calcEngine, string[] targets, JsonObject? windowConfiguration )
	{
		InitializeComponent();
		this.calcEngine = calcEngine.ToLower();
		this.windowConfiguration = windowConfiguration ?? new JsonObject();
		if ( this.windowConfiguration[ "calcEngines" ] == null )
		{
			this.windowConfiguration[ "calcEngines" ] = new JsonObject();
		}

		var calcEngineConfig = this.windowConfiguration[ "calcEngines" ]![ this.calcEngine ];

		dataSource.Items.Clear();
		dataSource.DisplayMember = "Value";
		dataSource.ValueMember = "Key";

		dataSource.DataSource =
			targets
				.Select( t => new KeyValuePair<string, string>( t, $"{t} Server" ) )
				.Concat( new KeyValuePair<string, string>[] {
					new( LocalProfileFile, "Downloaded Data Source File" ),
					new( LocalGroupFile, "Exported xDS File" )
				} )
				.ToArray();

		dataSource.SelectedValue = (string?)calcEngineConfig?[ nameof( dataSource ) ] ?? targets.First();

		this.clientName.Text =
			(string?)calcEngineConfig?[ nameof( clientName ) ] ??
			clientName ??
			Path.GetFileNameWithoutExtension( calcEngine ).Split( '_' ).First( p => !new[] { "Conduent", "Buck", "BTR" }.Contains( p ) );

		this.authId.Text = authId;

		loadTables.Checked = (bool?)calcEngineConfig?[ nameof( loadTables ) ] ?? true;
		downloadGlobalTables.Checked = !File.Exists( Path.Combine( AddIn.ResourcesPath, Constants.FileNames.GlobalTables ) );

		var defaultPath = Path.Combine( AddIn.ResourcesPath, "ConfigLookups", $"{this.clientName.Text}.xml" );
		var path = (string?)calcEngineConfig?[ nameof( configLookupPath ) ] ?? ( File.Exists( defaultPath ) ? defaultPath : null );
		downloadConfigLookups.Checked = string.IsNullOrEmpty( path ) || !File.Exists( path );
		configLookupPath.Text = path;

		configLookupUrl.Text =
			(string?)calcEngineConfig?[ nameof( configLookupUrl ) ] ??
			$"https://qabtr.lifeatworkportal.com/admin/{this.clientName.Text}/xml/config-lookups.xml";
	}

	bool IsLocalFile => new[] { LocalProfileFile, LocalGroupFile }.Contains( (string?)dataSource.SelectedValue );

	public LoadInputTabInfo? GetInfo( string? userName, string? password, NativeWindow? owner = null )
	{
		dataSource.Select();
		emailAddress.Text = userName;
		this.password.Text = password;

		DataSource_SelectedIndexChanged( this, EventArgs.Empty );
		LoadTables_CheckedChanged( this, EventArgs.Empty );

		var dialogResult = ShowDialog( owner );

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

		var calcEngines = ( windowConfiguration[ "calcEngines" ] as JsonObject )!;
		var calcEngineConfig = calcEngines[ calcEngine ];
		if ( calcEngineConfig != null )
		{
			calcEngines.Remove( calcEngine );
		}

		var dataSourceValue = (string)dataSource.SelectedValue!;

		calcEngines[ calcEngine ] = new JsonObject
		{
			[ nameof( clientName ) ] = clientName.Text,
			[ nameof( dataSource ) ] = dataSourceValue,
			[ nameof( loadTables ) ] = loadTables.Checked,
			[ nameof( configLookupPath ) ] = (string?)calcEngineConfig?[ nameof( configLookupPath ) ],
			[ nameof( configLookupUrl ) ] = (string?)calcEngineConfig?[ nameof( configLookupUrl ) ],
		};

		if ( dataSourceValue == LocalProfileFile || dataSourceValue == LocalGroupFile )
		{
			calcEngines[ calcEngine ]![ dataSourceValue.ToCamelCase() ] = inputFileName.Text;
		}
		if ( downloadConfigLookups.Checked )
		{
			calcEngines[ calcEngine ]![ nameof( configLookupUrl ) ] = configLookupUrl.Text;
		}
		else if ( !string.IsNullOrEmpty( configLookupPath.Text ) )
		{
			calcEngines[ calcEngine ]![ nameof( configLookupPath ) ] = configLookupPath.Text;
		}

		return new()
		{
			AuthId = dataSourceValue != LocalProfileFile ? authId.Text : null,
			ClientName = clientName.Text,
			DataSource = dataSourceValue,
			DataSourceFile = IsLocalFile ? inputFileName.Text : null,
			LoadLookupTables = loadTables.Checked,
			DownloadGlobalTables = downloadGlobalTables.Checked,
			ConfigLookupsUrl = downloadConfigLookups.Checked ? configLookupUrl.Text : null,
			ConfigLookupsPath = !downloadConfigLookups.Checked && !string.IsNullOrEmpty( configLookupPath.Text ) ? configLookupPath.Text : null,
			UserName = emailAddress.Text,
			Password = this.password.Text,
			WindowConfiguration = windowConfiguration
		};
	}

	private void LoadInputTab_Load( object sender, EventArgs e )
	{
		WindowState = Enum.TryParse( (string?)windowConfiguration[ "state" ], out FormWindowState state ) ? state : FormWindowState.Normal;
		Location = new Point { X = (int?)windowConfiguration[ "left" ] ?? Left, Y = (int?)windowConfiguration[ "top" ] ?? Top };
		Size = new Size { Width = (int?)windowConfiguration[ "width" ] ?? Width, Height = (int?)windowConfiguration[ "height" ] ?? Height };
	}

	private void LoadTables_CheckedChanged( object sender, EventArgs e )
	{
		downloadGlobalTables.Enabled = downloadConfigLookups.Enabled = loadTables.Checked;
		DownloadConfigLookups_CheckedChanged( sender, e );
	}

	private void DownloadConfigLookups_CheckedChanged( object sender, EventArgs e )
	{
		configLookupsUrlLabel.Enabled = configLookupUrl.Enabled = loadTables.Checked && downloadConfigLookups.Checked;
		configLookupPathSelect.Enabled = configLookupPathLabel.Enabled = configLookupPath.Enabled = loadTables.Checked && !downloadConfigLookups.Checked;
	}

	private void DataSource_SelectedIndexChanged( object sender, EventArgs e )
	{
		var key = (string?)dataSource.SelectedValue;

		emailAddress.Enabled = password.Enabled = !IsLocalFile;
		inputFileNameLabel.Enabled = inputFileName.Enabled = inputFileNameSelect.Enabled = IsLocalFile;
		authIdLabel.Enabled = authId.Enabled = key != LocalProfileFile;

		if ( IsLocalFile && !hasChangedInputFile )
		{
			var defaultInputFile = (string?)windowConfiguration[ "calcEngines" ]![ calcEngine ]?[ key!.ToCamelCase() ];
			inputFileName.Text = defaultInputFile;
		}
	}

	private void InputFileNameSelect_Click( object sender, EventArgs e )
	{
		var openDialog = new OpenFileDialog()
		{
			Filter = "Xml Files|*.xml|Json Files|*.json",
			Title = "Local Xml/Json Data",
			CheckFileExists = true,
			FileName = inputFileName.Text,
			RestoreDirectory = true,
			InitialDirectory = !string.IsNullOrEmpty( inputFileName.Text ) ? Path.GetDirectoryName( inputFileName.Text ) : null
		};

		if ( openDialog.ShowDialog() == DialogResult.OK )
		{
			inputFileName.Text = openDialog.FileName;
			hasChangedInputFile = true;
		}
	}

	private void ConfigLookupPathSelect_Click( object sender, EventArgs e )
	{
		var openDialog = new OpenFileDialog()
		{
			Filter = "Xml Files|*.xml",
			Title = "Config-Lookup Xml File",
			CheckFileExists = true,
			FileName = inputFileName.Text,
			RestoreDirectory = true,
			InitialDirectory = !string.IsNullOrEmpty( configLookupPath.Text ) ? Path.GetDirectoryName( configLookupPath.Text ) : null
		};

		if ( openDialog.ShowDialog() == DialogResult.OK )
		{
			configLookupPath.Text = openDialog.FileName;
		}
	}

	private void Ok_Click( object sender, EventArgs e )
	{
		errorProvider.Clear();

		if ( string.IsNullOrEmpty( clientName.Text ) )
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
		if ( dataSource.SelectedItem == null )
		{
			errorProvider.SetError( dataSource, "You must select at least one Data Source to continue." );
		}
		if ( (string?)dataSource.SelectedValue != LocalProfileFile && string.IsNullOrEmpty( authId.Text ) )
		{
			errorProvider.SetError( authId, "You must provide an Auth ID to continue." );
		}
		/*
		if ( string.IsNullOrEmpty( configLookupUrl.Text ) )
		{
			errorProvider.SetError( authId, "You must provide an Auth ID to continue." );
		}
		*/

		if ( loadTables.Checked )
		{
			if ( downloadConfigLookups.Checked && string.IsNullOrEmpty( configLookupUrl.Text ) )
			{
				errorProvider.SetError( configLookupUrl, "You must provide a Config Lookups URL to continue." );
			}
			else if ( !downloadConfigLookups.Checked && ( string.IsNullOrEmpty( configLookupPath.Text ) || !File.Exists( configLookupPath.Text ) ) )
			{
				errorProvider.SetError( configLookupPath, "You must provide a Config Lookups Path to existing file continue." );
			}
		}

		if ( !errorProvider.HasErrors )
		{
			DialogResult = DialogResult.OK;
			Close();
		}
	}
}