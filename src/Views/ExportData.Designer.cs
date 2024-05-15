namespace KAT.Camelot.Extensibility.Excel.AddIn;

partial class ExportData
{
	/// <summary>
	/// Required designer variable.
	/// </summary>
	private System.ComponentModel.IContainer components = null;

	/// <summary>
	/// Clean up any resources being used.
	/// </summary>
	/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
	protected override void Dispose( bool disposing )
	{
		if ( disposing && ( components != null ) )
		{
			components.Dispose();
		}
		base.Dispose( disposing );
	}

	#region Windows Form Designer generated code

	/// <summary>
	/// Required method for Designer support - do not modify
	/// the contents of this method with the code editor.
	/// </summary>
	private void InitializeComponent()
	{
		components = new System.ComponentModel.Container();
		cancel = new Button();
		ok = new Button();
		outputFileNameSelect = new Button();
		outputFileName = new TextBox();
		label3 = new Label();
		errorProvider = new ErrorProvider( components );
		clientName = new TextBox();
		lClientName = new Label();
		authIdsToExport = new TextBox();
		lAuthIdsToExport = new Label();
		limitFile = new CheckBox();
		limitFileSize = new TextBox();
		validateConfiguration = new Button();
		( (System.ComponentModel.ISupportInitialize)errorProvider ).BeginInit();
		SuspendLayout();
		// 
		// cancel
		// 
		cancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		cancel.DialogResult = DialogResult.Cancel;
		cancel.Location = new Point( 395, 189 );
		cancel.Margin = new Padding( 4, 3, 4, 3 );
		cancel.Name = "cancel";
		cancel.Size = new Size( 88, 27 );
		cancel.TabIndex = 9;
		cancel.Text = "Cancel";
		cancel.UseVisualStyleBackColor = true;
		// 
		// ok
		// 
		ok.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		ok.Location = new Point( 301, 189 );
		ok.Margin = new Padding( 4, 3, 4, 3 );
		ok.Name = "ok";
		ok.Size = new Size( 88, 27 );
		ok.TabIndex = 8;
		ok.Text = "&Export Data";
		ok.UseVisualStyleBackColor = true;
		ok.Click += Ok_Click;
		// 
		// outputFileNameSelect
		// 
		outputFileNameSelect.Anchor = AnchorStyles.Top | AnchorStyles.Right;
		outputFileNameSelect.Location = new Point( 442, 27 );
		outputFileNameSelect.Margin = new Padding( 4, 3, 4, 3 );
		outputFileNameSelect.Name = "outputFileNameSelect";
		outputFileNameSelect.Size = new Size( 40, 23 );
		outputFileNameSelect.TabIndex = 7;
		outputFileNameSelect.Text = "...";
		outputFileNameSelect.UseVisualStyleBackColor = true;
		outputFileNameSelect.Click += OutputFileNameSelect_Click;
		// 
		// outputFileName
		// 
		outputFileName.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		outputFileName.Location = new Point( 13, 27 );
		outputFileName.Margin = new Padding( 4, 3, 4, 3 );
		outputFileName.Name = "outputFileName";
		outputFileName.Size = new Size( 421, 23 );
		outputFileName.TabIndex = 6;
		// 
		// label3
		// 
		label3.AutoSize = true;
		label3.Location = new Point( 13, 9 );
		label3.Margin = new Padding( 4, 0, 4, 0 );
		label3.Name = "label3";
		label3.Size = new Size( 69, 15 );
		label3.TabIndex = 5;
		label3.Text = "&Output File:";
		// 
		// errorProvider
		// 
		errorProvider.BlinkStyle = ErrorBlinkStyle.NeverBlink;
		errorProvider.ContainerControl = this;
		// 
		// clientName
		// 
		clientName.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		clientName.Location = new Point( 13, 71 );
		clientName.Margin = new Padding( 4, 3, 4, 3 );
		clientName.Name = "clientName";
		clientName.Size = new Size( 469, 23 );
		clientName.TabIndex = 1;
		// 
		// lClientName
		// 
		lClientName.AutoSize = true;
		lClientName.Location = new Point( 13, 53 );
		lClientName.Margin = new Padding( 4, 0, 4, 0 );
		lClientName.Name = "lClientName";
		lClientName.Size = new Size( 73, 15 );
		lClientName.TabIndex = 0;
		lClientName.Text = "&Client Name";
		// 
		// authIdsToExport
		// 
		authIdsToExport.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		authIdsToExport.Location = new Point( 13, 115 );
		authIdsToExport.Margin = new Padding( 4, 3, 4, 3 );
		authIdsToExport.Name = "authIdsToExport";
		authIdsToExport.Size = new Size( 469, 23 );
		authIdsToExport.TabIndex = 11;
		// 
		// lAuthIdsToExport
		// 
		lAuthIdsToExport.AutoSize = true;
		lAuthIdsToExport.Location = new Point( 13, 97 );
		lAuthIdsToExport.Margin = new Padding( 4, 0, 4, 0 );
		lAuthIdsToExport.Name = "lAuthIdsToExport";
		lAuthIdsToExport.Size = new Size( 153, 15 );
		lAuthIdsToExport.TabIndex = 10;
		lAuthIdsToExport.Text = "&AuthID To Export (Optional)";
		// 
		// limitFile
		// 
		limitFile.AutoSize = true;
		limitFile.Location = new Point( 13, 145 );
		limitFile.Margin = new Padding( 4 );
		limitFile.Name = "limitFile";
		limitFile.Size = new Size( 247, 19 );
		limitFile.TabIndex = 12;
		limitFile.Text = "&Limit exported files to specified size in MB";
		limitFile.UseVisualStyleBackColor = true;
		limitFile.CheckedChanged += LimitFile_CheckedChanged;
		// 
		// limitFileSize
		// 
		limitFileSize.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		limitFileSize.Location = new Point( 298, 143 );
		limitFileSize.Margin = new Padding( 4, 3, 4, 3 );
		limitFileSize.Name = "limitFileSize";
		limitFileSize.Size = new Size( 185, 23 );
		limitFileSize.TabIndex = 13;
		// 
		// validateConfiguration
		// 
		validateConfiguration.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		validateConfiguration.Location = new Point( 140, 189 );
		validateConfiguration.Margin = new Padding( 4, 3, 4, 3 );
		validateConfiguration.Name = "validateConfiguration";
		validateConfiguration.Size = new Size( 153, 27 );
		validateConfiguration.TabIndex = 14;
		validateConfiguration.Text = "&Validate Configuration";
		validateConfiguration.UseVisualStyleBackColor = true;
		validateConfiguration.DialogResult = DialogResult.Retry;
		validateConfiguration.Click += ValidateConfiguration_Click;
		// 
		// ExportData
		// 
		AcceptButton = ok;
		AutoScaleDimensions = new SizeF( 7F, 15F );
		AutoScaleMode = AutoScaleMode.Font;
		CancelButton = cancel;
		ClientSize = new Size( 500, 228 );
		Controls.Add( validateConfiguration );
		Controls.Add( limitFileSize );
		Controls.Add( limitFile );
		Controls.Add( authIdsToExport );
		Controls.Add( lAuthIdsToExport );
		Controls.Add( clientName );
		Controls.Add( lClientName );
		Controls.Add( outputFileNameSelect );
		Controls.Add( outputFileName );
		Controls.Add( label3 );
		Controls.Add( ok );
		Controls.Add( cancel );
		Margin = new Padding( 4, 3, 4, 3 );
		MaximizeBox = false;
		MaximumSize = new Size( 814, 348 );
		MinimizeBox = false;
		MinimumSize = new Size( 371, 248 );
		Name = "ExportData";
		ShowIcon = false;
		ShowInTaskbar = false;
		Text = "Export Data";
		Load += ExportData_Load;
		( (System.ComponentModel.ISupportInitialize)errorProvider ).EndInit();
		ResumeLayout( false );
		PerformLayout();
	}

	#endregion
	private System.Windows.Forms.Button cancel;
	private System.Windows.Forms.Button ok;
	private System.Windows.Forms.Button outputFileNameSelect;
	private System.Windows.Forms.TextBox outputFileName;
	private System.Windows.Forms.Label label3;
	private System.Windows.Forms.ErrorProvider errorProvider;
	private System.Windows.Forms.TextBox clientName;
	private System.Windows.Forms.Label lClientName;
	private TextBox authIdsToExport;
	private Label lAuthIdsToExport;
	private TextBox limitFileSize;
	private CheckBox limitFile;
	private Button validateConfiguration;
}