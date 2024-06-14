using System.Text.Json.Nodes;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal delegate Task ProcessAsyncHandler( IProgress<int> currentProgress, IProgress<string> currentStatus, CancellationToken cancellationToken );

internal partial class Processing : Form
{
	private readonly JsonObject windowConfiguration;
	private readonly CancellationTokenSource cancellationTokenSource = new ();
	private int currentProgressValue;

	public Processing( string title, int maximum, JsonObject? windowConfiguration )
	{
		InitializeComponent();
		Text = title;
		this.windowConfiguration = windowConfiguration ?? new JsonObject();
		progressBar.Maximum = maximum;
	}

	private void Processing_Load( object sender, EventArgs e )
	{
		Location = new Point { X = (int?)windowConfiguration[ "left" ] ?? Left, Y = (int?)windowConfiguration[ "top" ] ?? Top };
		Size = new Size { Width = (int?)windowConfiguration[ "width" ] ?? Width, Height = (int?)windowConfiguration[ "height" ] ?? Height };
	}

	// http://stackoverflow.com/a/18033198/166231
	public async Task<ProcessingInfo> ProcessAsync( ProcessAsyncHandler action )
	{
		var result = DialogResult.None;

		Show();

		try
		{
			var currentProgress = new Progress<int>( i => progressBar.Value = currentProgressValue = Math.Min( progressBar.Maximum, i == -1 ? currentProgressValue + 1 : i ) );
			var currentStatus = new Progress<string>( s => processingLabel.Text = s );

			await action( currentProgress, currentStatus, cancellationTokenSource.Token );

			result = DialogResult.OK;
		}
		catch ( OperationCanceledException )
		{
			result = DialogResult.Cancel;
		}
		catch ( AggregateException ex )
		{
			// TODO: Do I ever hit this?  If so, look at Evo for 'unwrap' processing
			MessageBox.Show( ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error );
		}
		catch ( Exception ex )
		{
			ExcelDna.Logging.LogDisplay.RecordLine( $"Process Exception: {ex.Message}" );
			ExcelDna.Logging.LogDisplay.RecordLine( $"StackTrace: {ex.StackTrace}" );
		}

		Close();

		windowConfiguration[ "top" ] = Location.Y;
		windowConfiguration[ "left" ] = Location.X;
		windowConfiguration[ "height" ] = Size.Height;
		windowConfiguration[ "width" ] = Size.Width;

		return new() { Result = result, WindowConfiguration = windowConfiguration };
	}

	private void Cancel_Click( object sender, EventArgs e )
	{
		cancel.Enabled = false;
		cancellationTokenSource.Cancel();
	}
}
