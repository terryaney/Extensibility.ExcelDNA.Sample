using System.Text.Json.Nodes;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal partial class ExportData : Form
{
	private readonly JsonObject windowConfiguration;
	private bool isXml;

	public ExportData( JsonObject? windowConfiguration )
	{
		InitializeComponent();
		this.windowConfiguration = windowConfiguration ?? new JsonObject();
	}

	private void ExportData_Load( object sender, EventArgs e )
	{
		Location = new Point { X = (int?)windowConfiguration[ "left" ] ?? Left, Y = (int?)windowConfiguration[ "top" ] ?? Top };
	}

	public ExportDataInfo? GetInfo( string? clientName, string? outputFile, bool appendDateToName, bool isXml, NativeWindow owner )
	{
		this.clientName.Text = clientName;
		outputFileName.Text = outputFile;
		this.isXml = isXml;
		limitFileSize.Text = "10";
		limitFileSize.Enabled = false;

		if ( isXml )
		{
			var start = limitFileSize.Location.Y;

			lClientName.Visible = this.clientName.Visible = lAuthIdsToExport.Visible = authIdsToExport.Visible = false;

			limitFile.Location = new Point { X = limitFile.Location.X, Y = lClientName.Location.Y + 4 };
			limitFileSize.Location = new Point { X = limitFileSize.Location.X, Y = limitFile.Location.Y - 2 };

			var end = limitFileSize.Location.Y;
			var diff = start - end;

			Size = MinimumSize = MaximumSize = new Size { Width = Size.Width, Height = Size.Height - diff };
		}

		var dialogResult = ShowDialog( owner );

		if ( dialogResult == DialogResult.Cancel )
		{
			return null;
		}

		windowConfiguration[ "top" ] = Location.Y;
		windowConfiguration[ "left" ] = Location.X;

		return new()
		{
			WindowConfiguration = windowConfiguration,

			Action = dialogResult == DialogResult.OK ? ExportDataAction.Export : ExportDataAction.Validate,
			OutputFile = appendDateToName
				? Path.Combine( Path.GetDirectoryName( outputFileName.Text )!, Path.GetFileNameWithoutExtension( outputFileName.Text ) + "." + DateTime.Now.ToString( "yyyy.MM.dd" ) + ".xml" )
				: Path.Combine( Path.GetDirectoryName( outputFileName.Text )!, Path.GetFileNameWithoutExtension( outputFileName.Text ) + ".xml" ),
			ClientName = this.clientName.Text,
			AuthIdToExport = authIdsToExport.Text,
			MaxFileSize = limitFile.Checked ? int.Parse( limitFileSize.Text ) : null
		};
	}

	private void ValidateConfiguration_Click( object sender, EventArgs e )
	{
		DialogResult = DialogResult.Retry;
		Close();
	}

	private void LimitFile_CheckedChanged( object sender, EventArgs e )
	{
		limitFileSize.Enabled = limitFile.Checked;
	}

	private void Ok_Click( object sender, EventArgs e )
	{
		var isValid = true;
		errorProvider.Clear();

		if ( string.IsNullOrEmpty( outputFileName.Text ) )
		{
			errorProvider.SetError( outputFileNameSelect, "You must provide an output filename to continue." );
			isValid = false;
		}

		if ( !isXml && string.IsNullOrEmpty( clientName.Text ) )
		{
			errorProvider.SetError( clientName, "You must provide a client name to continue." );
			isValid = false;
		}

		if ( limitFile.Checked && ( string.IsNullOrEmpty( limitFileSize.Text ) || !int.TryParse( limitFileSize.Text, out var _ ) ) )
		{
			errorProvider.SetError( limitFileSize, "You must provide a valid integer file size limit to continue." );
			isValid = false;
		}

		if ( !isValid )
		{
			return;
		}

		DialogResult = DialogResult.OK;
		Close();
	}

	private void OutputFileNameSelect_Click( object sender, EventArgs e )
	{
		var saveDialog = new SaveFileDialog()
		{
			Filter = isXml ? "Xml Files|*.xml" : "Json Files|*.json",
			Title = isXml ? "Save Xml Data" : "Save Json Data",
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