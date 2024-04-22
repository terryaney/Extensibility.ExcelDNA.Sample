using System.Text.Json.Nodes;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Extensibility.Excel.AddIn;

internal partial class SaveHistory : Form
{
	private string currentVersion = null!;
	private readonly MSExcel.Workbook workbook;
	private readonly WorkbookState workbookState;
	private readonly JsonObject windowConfiguration;

	public SaveHistory( MSExcel.Workbook workbook, WorkbookState workbookState, JsonObject? windowConfiguration )
	{
		InitializeComponent();
		this.workbook = workbook;
		this.workbookState = workbookState;
		this.windowConfiguration = windowConfiguration ?? new JsonObject();
	}

	/// <summary>
	/// If an Excel file has 'RBLInfo' history log, this prompts the user to enter information about the changes they are making and allows uploading to Management Site if it is a CalcEngine.
	/// </summary>
	/// <returns>Returns action to perform when saving Excel file that has 'RBLInfo' history log.  Ignore - do nothing, simply allow save to occur.  OK - Update history log.  Continue - Update history log and upload to Management Site.  Retry - Do not update history log, just attempt to re-upload to Management Site.</returns>
	public SaveHistoryInfo GetHistoryInformation( string name, string? userName, string? password )
	{
		var sheets = workbook.Worksheets.Cast<MSExcel.Worksheet>();
		var historySheet =
			sheets.FirstOrDefault( s => workbook.Name == "RBL.Template.xlsx" && s.Name == "Version History" ) ??
			sheets.FirstOrDefault( s => s.Name == "Plan Info" /* && s.Names.Cast<MSExcel.Name>().Any( n => n.Name.EndsWith( "!SheetType" ) && (string)s.Range[ "SheetType" ].Text == "Plan Info" ) */ ) ??
			sheets.FirstOrDefault( s => s.Name == "RBLInfo" && s.Names.Cast<MSExcel.Name>().Any( n => n.Name.EndsWith( "!SheetType" ) && ( (string)s.Range[ "SheetType" ].Text ).StartsWith( "RBL Framework SpreadEngine" ) ) );

		var historyNames = historySheet?.Names.Cast<MSExcel.Name>();
		var historyRange = historyNames?.Where( n => n.Name.EndsWith( "!VersionHistory" ) ).Select( n => historySheet!.Range[ n.Name ] ).FirstOrDefault();

		if ( historySheet == null || historyNames == null || historyRange == null )
		{
			return new() { Result = DialogResult.Ignore, WindowConfiguration = windowConfiguration };
		}

		var currentVersionRange =
			historyNames.Where( n => n.Name.EndsWith( "!CurrentVersion" ) ).Select( n => historySheet.Range[ n.Name ] ).FirstOrDefault() ??
			historyNames.Where( n => n.Name.EndsWith( "!Version" ) ).Select( n => historySheet.Range[ n.Name ] ).FirstOrDefault() ??
			workbook.RangeOrNull<MSExcel.Range>( "Version" );

		if ( currentVersionRange == null )
		{
			return new() { Result = DialogResult.Ignore, WindowConfiguration = windowConfiguration };
		}

		string proposedVersion = null!;
		var currentVersionValue = (double?)currentVersionRange.Value;
		currentVersion = (string)currentVersionRange.Text;

		if ( currentVersionValue == null )
		{
			currentVersionRange.Value = 1;
			// Just get a 'formatted' version number
			proposedVersion = (string)currentVersionRange.Text;
			currentVersionRange.Value = null;
		}
		else
		{
			var versionParts = currentVersion.Split( '.' );
			var decimals = versionParts.Length == 2 ? versionParts[ 1 ].Length : 0;
			var currentFloor = Math.Floor( currentVersionValue.Value );

			currentVersionRange.Value = currentFloor + ( ( ( currentVersionValue - currentFloor ) * Math.Pow( 10, decimals ) ) + 1 ) / Math.Pow( 10, decimals );
			// Just get a 'formatted' version number
			proposedVersion = (string)currentVersionRange.Text;
			currentVersionRange.Value = currentVersionValue;
		}

		author.Text = name;
		version.Text = proposedVersion;
		lManagementSite.Text += $" (Current Version: {workbookState.UploadedVersion})";

		tUserName.Text = userName;
		tPassword.Text = password;

		if ( !workbookState.ShowCalcEngineManagement || !workbookState.IsCalcEngine )
		{
			ok.Text = "A&pply";
			lManagementSite.Visible = lUserName.Visible = lPassword.Visible = tUserName.Visible = tPassword.Visible = forceUpload.Visible = false;
			description.Height = ok.Top - description.Top - 20;
			MinimumSize = new Size( MinimumSize.Width, 235 );
			Height = 235;
		}
		else if ( string.Compare( workbookState.CheckedOutBy, userName, true ) != 0 )
		{
			ok.Text = "A&pply";
			lManagementSite.Text = "Check Out To Upload To Management Site";
			tUserName.Enabled = tPassword.Enabled = forceUpload.Enabled = false;
		}

		if ( currentVersionValue != null && tUserName.Visible && tUserName.Enabled )
		{
			versionLabel.Text = $"&Version (Use {currentVersion} to skip history update):";
		}

		description.Select();

		WindowState = Enum.TryParse( (string?)windowConfiguration[ "state" ], out FormWindowState state) ? state : FormWindowState.Normal;
		Location = new Point { X = (int?)windowConfiguration[ "left" ] ?? Left, Y = (int?)windowConfiguration[ "top" ] ?? Top };
		Size = new Size { Width = (int?)windowConfiguration[ "width" ] ?? Width, Height = (int?)windowConfiguration[ "height" ] ?? Height };

		var dialogResult = ShowDialog();

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
			Result = version.Text == currentVersion && dialogResult == DialogResult.Continue
				? DialogResult.Retry
				: dialogResult,
			Author = author.Text,
			Version = version.Text,
			Description = description.Text,
			UserName = tUserName.Text,
			Password = tPassword.Text,
			ForceUpload = forceUpload.Checked,

			HistoryRange = historyRange,
			VersionRange = currentVersionRange,

			WindowConfiguration = windowConfiguration
		};
	}

	private void Ok_Click( object sender, EventArgs e )
	{
		errorProvider.Clear();

		var uploadCalcEngine = tUserName.Visible && tUserName.Enabled && !string.IsNullOrEmpty( tUserName.Text ) && !string.IsNullOrEmpty( tPassword.Text );
		var isValid = IsValid();

		if ( isValid )
		{
			DialogResult = uploadCalcEngine ? DialogResult.Continue : DialogResult.OK;
			Close();
		}
	}

	private bool IsValid()
	{
		var isValid = true;

		if ( string.IsNullOrEmpty( author.Text ) )
		{
			errorProvider.SetError( author, "You must provide an Author to continue." );
			isValid = false;
		}
		if ( string.IsNullOrEmpty( version.Text ) )
		{
			errorProvider.SetError( version, "You must provide a Version to continue." );
			isValid = false;
		}
		else if ( !double.TryParse( version.Text, out var result ) )
		{
			errorProvider.SetError( version, "You must supply a numeric value for the Version number." );
			isValid = false;
		}
		else if ( double.TryParse( currentVersion, out var previous ) && result <= previous )
		{
			errorProvider.SetError( version, $"You must supply version number greater than the current version of {currentVersion}." );
			isValid = false;
		}

		return isValid;
	}
}