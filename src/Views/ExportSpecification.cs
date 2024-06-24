using System.Text.Json.Nodes;
using KAT.Camelot.Domain.Extensions;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal partial class ExportSpecification : Form
{
	private readonly string configKey;
	private readonly string saveSpecificationLocation;
	private readonly JsonObject windowConfiguration;

	public ExportSpecification( string saveSpecificationLocation, bool saveSpecification, JsonObject? windowConfiguration )
	{
		InitializeComponent();

		configKey = Path.GetFileName( saveSpecificationLocation ).ToLower();

		this.windowConfiguration = windowConfiguration ?? new JsonObject();
		if ( this.windowConfiguration[ "files" ] == null )
		{
			this.windowConfiguration[ "files" ] = new JsonObject();
		}

		var configFile = this.windowConfiguration[ "files" ]![ configKey ];
		var targets = 
 			( configFile?[ "locations" ] as JsonArray )?
				.Select( l => new SpecificationLocation { Location = (string)l![ "location" ]!, Selected = (bool)l[ "selected" ]! } ) ?? Enumerable.Empty<SpecificationLocation>();

		this.locations.Items.Clear();
		this.locations.Items.AddRange( targets.Select( t => t.Location ).ToArray() );

		foreach( var target in targets.Where( t => t.Selected ) )
		{
			this.locations.SetItemChecked( this.locations.Items.IndexOf( target.Location ), true );
		}

		this.saveSpecification.Checked = (bool?)configFile?[ "save" ] ?? saveSpecification;
		
		var locationParts = saveSpecificationLocation.Split( '\\' );
		this.saveSpecification.Text = $"&Save Specification file to ..\\{string.Join( "\\", locationParts.Skip( locationParts.Length - 2 ) )}";

		var toolTip = new ToolTip();
		toolTip.SetToolTip( this.saveSpecification, $"Save Specification file to {saveSpecificationLocation}" );

		this.saveSpecificationLocation = saveSpecificationLocation;
	}

	public ExportSpecificationInfo? GetInfo( NativeWindow? owner = null )
	{
		ok.Select();

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

		var configFiles = ( windowConfiguration[ "files" ] as JsonObject )!;
		var configFile = configFiles[ configKey ];
		if ( configFile != null )
		{
			configFiles.Remove( configKey );
		}

		configFiles[ configKey ] = new JsonObject
		{
			[ "save" ] = saveSpecification.Checked,
			[ "locations" ] = locations.Items.Cast<string>().Select( l => new JsonObject { [ "location" ] = l, [ "selected" ] = locations.CheckedItems.Contains( l ) } ).ToJsonArray()
		};

		return new()
		{
			SaveSpecification = saveSpecification.Checked,
			Locations = locations.Items.Cast<string>().Select( l => new SpecificationLocation { Location = l, Selected = locations.CheckedItems.Contains( l ) } ).ToArray(),
			WindowConfiguration = windowConfiguration
		};
	}

	private void ExportSpecification_Load( object sender, EventArgs e )
	{
		WindowState = Enum.TryParse( (string?)windowConfiguration[ "state" ], out FormWindowState state ) ? state : FormWindowState.Normal;
		Location = new Point { X = (int?)windowConfiguration[ "left" ] ?? Left, Y = (int?)windowConfiguration[ "top" ] ?? Top };
		Size = new Size { Width = (int?)windowConfiguration[ "width" ] ?? Width, Height = (int?)windowConfiguration[ "height" ] ?? Height };
	}

	private void Ok_Click( object sender, EventArgs e )
	{
		errorProvider.Clear();

		if ( locations.CheckedItems.Count == 0 )
		{
			errorProvider.SetError( locations, "You must select at least one export location to continue." );
		}

		if ( !errorProvider.HasErrors )
		{
			DialogResult = DialogResult.OK;
			Close();
		}
	}

	private void Locations_KeyDown( object sender, KeyEventArgs e )
	{
		if ( e.KeyCode == Keys.Delete )
		{
			// Delete key was pressed
			// Check if an item is selected in the list
			if ( locations.SelectedItem != null )
			{
				// Remove the selected item
				locations.Items.Remove( locations.SelectedItem );
			}
		}
	}

	private void AddLocation_Click( object sender, EventArgs e )
	{
		var specProjectFolder = new DirectoryInfo( Path.GetDirectoryName( saveSpecificationLocation )! ).Parent!.FullName;
		var defaultFileName = Path.Combine( specProjectFolder, "Xml" );

		if ( locations.Items.Cast<string>().FirstOrDefault( l => string.Compare( l, defaultFileName, true ) == 0 ) != null )
		{
			defaultFileName = null;
		}

		var openDialog = new FolderBrowserDialog()
		{
			AutoUpgradeEnabled = true,
			Description = "Select Export Location",
			UseDescriptionForTitle = true,
			ShowNewFolderButton = true,
			InitialDirectory = defaultFileName ?? ""
		};

		if ( openDialog.ShowDialog() == DialogResult.OK )
		{
			var item = locations.Items.Cast<string>().FirstOrDefault( l => string.Compare( l, openDialog.SelectedPath, true ) == 0 );
			if ( item != null )
			{
				locations.SetItemChecked( locations.Items.IndexOf( item ), true );
			}
			else
			{
				locations.Items.Add( openDialog.SelectedPath );
				locations.SetItemChecked( locations.Items.Count - 1, true );
			}
		}
	}
}