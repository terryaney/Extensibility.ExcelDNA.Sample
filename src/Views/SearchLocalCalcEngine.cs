using System.Text.Json.Nodes;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal partial class SearchLocalCalcEngines : Form
{
	private readonly JsonObject windowConfiguration;

	public SearchLocalCalcEngines( JsonObject? windowConfiguration )
	{
		InitializeComponent();
		this.windowConfiguration = windowConfiguration ?? new JsonObject();
	}

	public SearchLocalCalcEnginesInfo? Search()
	{
		searchLocation.Text = (string?)windowConfiguration[ "searchLocation" ] ?? string.Empty;
		tokensToFind.Text = (string?)windowConfiguration[ "tokensToFind" ] ?? string.Empty;

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
		windowConfiguration[ "searchLocation" ] = searchLocation.Text;
		windowConfiguration[ "tokensToFind" ] = tokensToFind.Text;

		return new()
		{
			Folder = searchLocation.Text,
			Tokens = tokensToFind.Text.Split( '|' ).Select( t => t.Trim() ).ToArray(),
			WindowConfiguration = windowConfiguration
		};
	}

	private void SearchLocalCalcEngines_Load( object sender, EventArgs e )
	{
		WindowState = Enum.TryParse( (string?)windowConfiguration[ "state" ], out FormWindowState state) ? state : FormWindowState.Normal;
		Location = new Point { X = (int?)windowConfiguration[ "left" ] ?? Left, Y = (int?)windowConfiguration[ "top" ] ?? Top };
		Size = new Size { Width = (int?)windowConfiguration[ "width" ] ?? Width, Height = (int?)windowConfiguration[ "height" ] ?? Height };
	}

	private void SelectFolder( TextBox target )
	{
		searchLocationDialog.InitialDirectory = target.Text;
		if ( searchLocationDialog.ShowDialog() == DialogResult.OK )
		{
			target.Text = searchLocationDialog.SelectedPath;
		}
	}

	private void SearchLocationSelect_Click( object sender, EventArgs e ) => SelectFolder( searchLocation );

	private void Ok_Click( object sender, EventArgs e )
	{
		DialogResult = DialogResult.OK;
		Close();
	}
}