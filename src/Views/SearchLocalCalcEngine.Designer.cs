namespace KAT.Camelot.Extensibility.Excel.AddIn;

partial class SearchLocalCalcEngines
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
		searchLocationLabel = new Label();
		searchLocation = new TextBox();
		tokensToFindLabel = new Label();
		tokensToFind = new TextBox();
		searchLocationDialog = new FolderBrowserDialog();
		searchLocationSelect = new Button();
		cancel = new Button();
		ok = new Button();
		SuspendLayout();
		// 
		// label1
		// 
		searchLocationLabel.AutoSize = true;
		searchLocationLabel.Location = new Point( 16, 14 );
		searchLocationLabel.Margin = new Padding( 4, 0, 4, 0 );
		searchLocationLabel.Name = "label1";
		searchLocationLabel.Size = new Size( 98, 20 );
		searchLocationLabel.TabIndex = 2;
		searchLocationLabel.Text = "&Folder Name:";
		// 
		// searchLocation
		// 
		searchLocation.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		searchLocation.Location = new Point( 20, 39 );
		searchLocation.Margin = new Padding( 4, 5, 4, 5 );
		searchLocation.Name = "searchLocation";
		searchLocation.Size = new Size( 431, 27 );
		searchLocation.TabIndex = 3;
		// 
		// searchLocationSelect
		// 
		searchLocationSelect.Anchor = AnchorStyles.Top | AnchorStyles.Right;
		searchLocationSelect.Location = new Point( 460, 39 );
		searchLocationSelect.Margin = new Padding( 4, 5, 4, 5 );
		searchLocationSelect.Name = "exportFileNameSelect";
		searchLocationSelect.Size = new Size( 45, 28 );
		searchLocationSelect.TabIndex = 4;
		searchLocationSelect.Text = "...";
		searchLocationSelect.UseVisualStyleBackColor = true;
		searchLocationSelect.Click += SearchLocationSelect_Click;
		// 
		// cancel
		// 
		cancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		cancel.DialogResult = DialogResult.Cancel;
		cancel.Location = new Point( 405, 137 );
		cancel.Margin = new Padding( 4, 5, 4, 5 );
		cancel.Name = "cancel";
		cancel.Size = new Size( 100, 35 );
		cancel.TabIndex = 1;
		cancel.Text = "Cancel";
		cancel.UseVisualStyleBackColor = true;
		// 
		// ok
		// 
		ok.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		ok.Location = new Point( 297, 137 );
		ok.Margin = new Padding( 4, 5, 4, 5 );
		ok.Name = "ok";
		ok.Size = new Size( 100, 35 );
		ok.TabIndex = 0;
		ok.Text = "OK";
		ok.UseVisualStyleBackColor = true;
		ok.Click += Ok_Click;
		// 
		// tokensToFindLabel
		// 
		tokensToFindLabel.AutoSize = true;
		tokensToFindLabel.Location = new Point(15, 71);
		tokensToFindLabel.Margin = new Padding(4, 0, 4, 0);
		tokensToFindLabel.Name = "tokensToFindLabel";
		tokensToFindLabel.Size = new Size(155, 20);
		tokensToFindLabel.TabIndex = 12;
		tokensToFindLabel.Text = "&Token(s) (| delimitted)";
		// 
		// tokensToFind
		// 
		tokensToFind.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		tokensToFind.Location = new Point( 19, 96 );
		tokensToFind.Margin = new Padding( 4, 5, 4, 5 );
		tokensToFind.Name = "tokensToFind";
		tokensToFind.Size = new Size( 486, 27 );
		tokensToFind.TabIndex = 13;
		// 
		// SearchLocalCalcEngine
		// 
		AcceptButton = ok;
		AutoScaleDimensions = new SizeF( 8F, 20F );
		AutoScaleMode = AutoScaleMode.Font;
		CancelButton = cancel;
		ClientSize = new Size( 521, 179 );
		Controls.Add( tokensToFind );
		Controls.Add( tokensToFindLabel );
		Controls.Add( ok );
		Controls.Add( cancel );
		Controls.Add( searchLocationSelect );
		Controls.Add( searchLocation );
		Controls.Add( searchLocationLabel );
		Margin = new Padding( 4, 5, 4, 5 );
		MaximizeBox = false;
		MaximumSize = new Size( 927, 313 );
		MinimizeBox = false;
		MinimumSize = new Size( 421, 226 );
		Name = "SearchLocalCalcEngine";
		ShowIcon = false;
		ShowInTaskbar = false;
		Text = "Search Local CalcEngines...";
		this.Load += new System.EventHandler(this.SearchLocalCalcEngines_Load);
		ResumeLayout( false );
		PerformLayout();
	}

	#endregion

	private System.Windows.Forms.Label searchLocationLabel;
	private System.Windows.Forms.TextBox searchLocation;
	private System.Windows.Forms.FolderBrowserDialog searchLocationDialog;
	private System.Windows.Forms.Button searchLocationSelect;
	private System.Windows.Forms.Button cancel;
	private System.Windows.Forms.Button ok;
	private System.Windows.Forms.Label tokensToFindLabel;
	private System.Windows.Forms.TextBox tokensToFind;
}