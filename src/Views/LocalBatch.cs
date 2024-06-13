using System.Text.Json.Nodes;
using KAT.Camelot.RBLe.Core.Calculations;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal partial class LocalBatch : Form
{
	private readonly string calcEngine;
	private readonly JsonObject windowConfiguration;
	private readonly CancellationTokenSource cancellationSource;

	public LocalBatch( CalcEngineConfiguration configuration, string? currentTab, JsonObject? windowConfiguration )
	{
		InitializeComponent();

		this.calcEngine = configuration.Name;
		this.windowConfiguration = windowConfiguration ?? new JsonObject();

		if ( this.windowConfiguration[ "calcEngines" ] == null )
		{
			this.windowConfiguration[ "calcEngines" ] = new JsonObject();
		}

		var calcEngineConfig = this.windowConfiguration[ "calcEngines" ]![ this.calcEngine ];

		exportType.Items.Clear();
		exportType.Items.AddRange( new[]
		{
			"Export first table to CSV",
			"Export first table to CSV (transposing rows into column headers)",
			"Export all results to Xml file"
		} );

		exportType.SelectedIndex = (int?)calcEngineConfig?[ nameof( exportType ) ] ?? 0;

		var toolTip = new ToolTip();
		toolTip.SetToolTip( filter, "Sample: HistoryData/HistoryItem[@hisType='Status'][position()=last()]/status='A'" );
		filter.Text = (string?)calcEngineConfig?[ nameof( filter ) ];

		inputFileName.Text = (string?)calcEngineConfig?[ nameof( inputFileName ) ];
		outputFileName.Text = (string?)calcEngineConfig?[ nameof( outputFileName ) ];

		inputTab.Items.Clear();
		inputTab.DisplayMember = "Value";
		inputTab.ValueMember = "Key";
		inputTab.DataSource =
			configuration.InputTabs
				.Select( t => new KeyValuePair<string, string>( t.Name, t.Name ) )
				.ToArray();

		inputTab.SelectedValue = 
			(string?)calcEngineConfig?[ nameof( inputTab ) ] ?? 
			configuration.InputTabs.Select( t => t.Name ).FirstOrDefault( n => n == currentTab ) ??
			configuration.InputTabs.Select( t => t.Name ).FirstOrDefault();

		resultTab.Items.Clear();
		resultTab.DisplayMember = "Value";
		resultTab.ValueMember = "Key";
		resultTab.DataSource =
			configuration.ResultTabs
				.Select( t => new KeyValuePair<string, string>( t.Name, t.Name ) )
				.ToArray();
		resultTab.SelectedValue = 
			(string?)calcEngineConfig?[ nameof( resultTab ) ] ?? 
			configuration.ResultTabs.Select( t => t.Name ).FirstOrDefault( n => n == currentTab ) ??
			configuration.ResultTabs.Select( t => t.Name ).FirstOrDefault();

		limitRows.Checked = (bool?)calcEngineConfig?[ nameof( limitRows ) ] ?? false;
		limitRowsTo.Text = (string?)calcEngineConfig?[ nameof( limitRowsTo ) ] ?? "5";

		saveErrorCalcEngineError.Checked = (bool?)calcEngineConfig?[ nameof( saveErrorCalcEngineError ) ] ?? false;
		saveErrorCalcEngineCount.Text = (string?)calcEngineConfig?[ nameof( saveErrorCalcEngineCount ) ] ?? "5";

		cancellationSource = new CancellationTokenSource();
	}

	private void LocalBatch_Load( object sender, EventArgs e )
	{
		Location = new Point { X = (int?)windowConfiguration[ "left" ] ?? Left, Y = (int?)windowConfiguration[ "top" ] ?? Top };
		Size = new Size { Width = (int?)windowConfiguration[ "width" ] ?? Width, Height = (int?)windowConfiguration[ "height" ] ?? Height };
	}

	public LocalBatchInfo? GetInfo( NativeWindow owner )
	{
		SaveErrorCalcEngineError_CheckedChanged( this, EventArgs.Empty );
		LimitRows_CheckedChanged( this, EventArgs.Empty );

		var dialogResult = ShowDialog( owner );

		if ( dialogResult != DialogResult.OK )
		{
			return null;
		}

		windowConfiguration[ "top" ] = Location.Y;
		windowConfiguration[ "left" ] = Location.X;
		windowConfiguration[ "height" ] = Size.Height;
		windowConfiguration[ "width" ] = Size.Width;

		var calcEngines = ( windowConfiguration[ "calcEngines" ] as JsonObject )!;
		var calcEngineConfig = calcEngines[ calcEngine ];
		if ( calcEngineConfig != null )
		{
			calcEngines.Remove( calcEngine );
		}

		calcEngines[ calcEngine ] = new JsonObject
		{
			[ nameof( exportType ) ] = exportType.SelectedIndex,
			[ nameof( filter ) ] = this.filter.Text,
			[ nameof( inputFileName ) ] = this.inputFileName.Text,
			[ nameof( outputFileName ) ] = this.outputFileName.Text,
			[ nameof( inputTab ) ] = (string)inputTab.SelectedValue!,
			[ nameof( resultTab ) ] = (string)resultTab.SelectedValue!,
			[ nameof( limitRows ) ] = limitRows.Checked,
			[ nameof( limitRowsTo ) ] = limitRowsTo.Text,
			[ nameof( saveErrorCalcEngineError ) ] = saveErrorCalcEngineError.Checked,
			[ nameof( saveErrorCalcEngineCount ) ] = saveErrorCalcEngineCount.Text
		};

		return new()
		{
			WindowConfiguration = windowConfiguration,

			InputFile = this.inputFileName.Text,
			OutputFile = this.outputFileName.Text,
			Filter = this.filter.Text,
			InputTab = (string)inputTab.SelectedValue!,
			ResultTab = (string)resultTab.SelectedValue!,
			ExportType = ExportFormat,
			InputRows = limitRows.Checked ? int.Parse( limitRowsTo.Text ) : null,
			ErrorCalcEngines = saveErrorCalcEngineError.Checked ? int.Parse( saveErrorCalcEngineCount.Text ) : null
		};
	}

	ExportFormatType ExportFormat { get { return (ExportFormatType)exportType.SelectedIndex; } }

	private void Ok_Click( object sender, EventArgs e )
	{
		var isValid = true;
		errorProvider.Clear();

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
		if ( inputTab.SelectedItem == null )
		{
			errorProvider.SetError( inputTab, "You must select an input tab to continue." );
			isValid = false;
		}
		if ( resultTab.SelectedItem == null )
		{
			errorProvider.SetError( resultTab, "You must select a result tab to continue." );
			isValid = false;
		}
		if ( exportType.SelectedItem == null )
		{
			errorProvider.SetError( exportType, "You must select an export type to continue." );
			isValid = false;
		}
		else if ( ExportFormat == ExportFormatType.Xml && string.Compare( Path.GetExtension( outputFileName.Text ), ".xml", true ) != 0 )
		{
			errorProvider.SetError( exportType, "You have selected an Xml export type, but the output filename does not end in .xml." );
			isValid = false;
		}
		else if ( ExportFormat != ExportFormatType.Xml && string.Compare( Path.GetExtension( outputFileName.Text ), ".csv", true ) != 0 )
		{
			errorProvider.SetError( exportType, "You have selected an Csv export type, but the output filename does not end in .csv." );
			isValid = false;
		}
		if ( limitRows.Checked && !int.TryParse( limitRowsTo.Text, out var _ ) )
		{
			errorProvider.SetError( limitRowsTo, "You must provide a valid number (no decimals) of data rows to process to continue." );
			isValid = false;
		}
		if ( saveErrorCalcEngineError.Checked && !int.TryParse( saveErrorCalcEngineCount.Text, out var _ ) )
		{
			errorProvider.SetError( saveErrorCalcEngineCount, "You must provide a valid number (no decimals) of possible error CalcEngines allowed to be saved." );
			isValid = false;
		}

		if ( !isValid )
		{
			return;
		}

		DialogResult = DialogResult.OK;
		Close();
	}

	private void LimitRows_CheckedChanged( object sender, EventArgs e )
	{
		if ( limitRowsTo.Enabled = limitRows.Checked )
		{
			limitRowsTo.Select();
		}
	}

	private void SaveErrorCalcEngineError_CheckedChanged( object sender, EventArgs e )
	{
		if ( saveErrorCalcEngineCount.Enabled = saveErrorCalcEngineError.Checked )
		{
			saveErrorCalcEngineCount.Select();
		}
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