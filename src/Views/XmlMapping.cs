using System.Text.Json.Nodes;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal partial class XmlMapping : Form
{
	private readonly JsonObject windowConfiguration;

	public XmlMapping( JsonObject? windowConfiguration )
	{
		InitializeComponent();
		this.windowConfiguration = windowConfiguration ?? new JsonObject();
	}

	private void XmlMapping_Load( object sender, EventArgs e )
	{
		Location = new Point { X = (int?)windowConfiguration[ "left" ] ?? Left, Y = (int?)windowConfiguration[ "top" ] ?? Top };
		Size = new Size { Width = (int?)windowConfiguration[ "width" ] ?? Width, Height = (int?)windowConfiguration[ "height" ] ?? Height };
	}

	public XmlMappingInfo? GetInfo( NativeWindow owner )
	{

		this.clientName.Text = (string?)windowConfiguration[ "clientName" ];
		this.inputFileName.Text = (string?)windowConfiguration[ "inputFile" ];
		this.outputFileName.Text = (string?)windowConfiguration[ "outputFile" ];
		
		var dialogResult = ShowDialog( owner );

		if ( dialogResult != DialogResult.OK )
		{
			return null;
		}

		windowConfiguration[ "clientName" ] = this.clientName.Text;
		windowConfiguration[ "inputFile" ] = this.inputFileName.Text;
		windowConfiguration[ "ouputFile" ] = this.outputFileName.Text;

		windowConfiguration[ "top" ] = Location.Y;
		windowConfiguration[ "left" ] = Location.X;
		windowConfiguration[ "height" ] = Size.Height;
		windowConfiguration[ "width" ] = Size.Width;

		return new()
		{
			WindowConfiguration = windowConfiguration,

			ClientName = this.clientName.Text,
			InputFile = this.inputFileName.Text,
			OutputFile = this.outputFileName.Text
		};
	}

	private void Ok_Click( object sender, EventArgs e )
	{
		var isValid = true;
		errorProvider.Clear();

		if ( string.IsNullOrEmpty( clientName.Text ) )
		{
			errorProvider.SetError( clientName, "You must provide a client name to continue." );
			isValid = false;
		}

		if ( string.IsNullOrEmpty( inputFileName.Text ) )
		{
			errorProvider.SetError( inputFileNameSelect, "You must provide an input filename to continue." );
			isValid = false;
		}
		else if ( !File.Exists( inputFileName.Text ) )
		{
			errorProvider.SetError( inputFileNameSelect, $"The selected file ({inputFileName.Text}) does not exist." );
			isValid = false;
		}

		if ( string.IsNullOrEmpty( outputFileName.Text ) )
		{
			errorProvider.SetError( outputFileNameSelect, "You must provide an output filename to continue." );
			isValid = false;
		}

		if ( !isValid )
		{
			return;
		}

		DialogResult = DialogResult.OK;
		Close();
	}

	private void InputFileNameSelect_Click( object sender, EventArgs e )
	{
		var openDialog = new OpenFileDialog()
		{
			Filter = "Xml Files|*.xml",
			Title = "Input Xml Data",
			CheckFileExists = true,
			FileName = inputFileName.Text,
			RestoreDirectory = true,
			InitialDirectory = !string.IsNullOrEmpty( inputFileName.Text ) ? Path.GetDirectoryName( inputFileName.Text ) : null
		};

		if ( openDialog.ShowDialog() == DialogResult.OK )
		{
			inputFileName.Text = openDialog.FileName;
		}
	}

	private void OuputFileNameSelect_Click( object sender, EventArgs e )
	{
		var saveDialog = new SaveFileDialog()
		{
			Filter = "Xml Files|*.xml|Csv Files|*.csv",
			Title = "Save Xml/Csv Result Data",
			OverwritePrompt = true,
			FileName = outputFileName.Text,
			RestoreDirectory = true,
			InitialDirectory = !string.IsNullOrEmpty( outputFileName.Text ) ? Path.GetDirectoryName( outputFileName.Text ) : null
		};

		if ( saveDialog.ShowDialog() == DialogResult.OK )
		{
			outputFileName.Text = saveDialog.FileName;
		}
	}
}