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
        // searchLocationLabel
        // 
        searchLocationLabel.AutoSize = true;
        searchLocationLabel.Location = new Point(14, 10);
        searchLocationLabel.Margin = new Padding(4, 0, 4, 0);
        searchLocationLabel.Name = "searchLocationLabel";
        searchLocationLabel.Size = new Size(78, 15);
        searchLocationLabel.TabIndex = 0;
        searchLocationLabel.Text = "&Folder Name:";
        // 
        // searchLocation
        // 
        searchLocation.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        searchLocation.Location = new Point(18, 29);
        searchLocation.Margin = new Padding(4, 4, 4, 4);
        searchLocation.Name = "searchLocation";
        searchLocation.Size = new Size(378, 23);
        searchLocation.TabIndex = 1;
        // 
        // tokensToFindLabel
        // 
        tokensToFindLabel.AutoSize = true;
        tokensToFindLabel.Location = new Point(13, 53);
        tokensToFindLabel.Margin = new Padding(4, 0, 4, 0);
        tokensToFindLabel.Name = "tokensToFindLabel";
        tokensToFindLabel.Size = new Size(122, 15);
        tokensToFindLabel.TabIndex = 3;
        tokensToFindLabel.Text = "&Token(s) (| delimitted)";
        // 
        // tokensToFind
        // 
        tokensToFind.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        tokensToFind.Location = new Point(17, 72);
        tokensToFind.Margin = new Padding(4, 4, 4, 4);
        tokensToFind.Name = "tokensToFind";
        tokensToFind.Size = new Size(426, 23);
        tokensToFind.TabIndex = 4;
        // 
        // searchLocationSelect
        // 
        searchLocationSelect.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        searchLocationSelect.Location = new Point(402, 29);
        searchLocationSelect.Margin = new Padding(4, 4, 4, 4);
        searchLocationSelect.Name = "searchLocationSelect";
        searchLocationSelect.Size = new Size(39, 21);
        searchLocationSelect.TabIndex = 2;
        searchLocationSelect.Text = "...";
        searchLocationSelect.UseVisualStyleBackColor = true;
        searchLocationSelect.Click += SearchLocationSelect_Click;
        // 
        // cancel
        // 
        cancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        cancel.DialogResult = DialogResult.Cancel;
        cancel.Location = new Point(354, 103);
        cancel.Margin = new Padding(4, 4, 4, 4);
        cancel.Name = "cancel";
        cancel.Size = new Size(88, 26);
        cancel.TabIndex = 6;
        cancel.Text = "Cancel";
        cancel.UseVisualStyleBackColor = true;
        // 
        // ok
        // 
        ok.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        ok.Location = new Point(260, 103);
        ok.Margin = new Padding(4, 4, 4, 4);
        ok.Name = "ok";
        ok.Size = new Size(88, 26);
        ok.TabIndex = 5;
        ok.Text = "OK";
        ok.UseVisualStyleBackColor = true;
        ok.Click += Ok_Click;
        // 
        // SearchLocalCalcEngines
        // 
        AcceptButton = ok;
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        CancelButton = cancel;
        ClientSize = new Size(456, 140);
        Controls.Add(tokensToFind);
        Controls.Add(tokensToFindLabel);
        Controls.Add(ok);
        Controls.Add(cancel);
        Controls.Add(searchLocationSelect);
        Controls.Add(searchLocation);
        Controls.Add(searchLocationLabel);
        Margin = new Padding(4, 4, 4, 4);
        MaximizeBox = false;
        MaximumSize = new Size(813, 245);
        MinimizeBox = false;
        MinimumSize = new Size(370, 179);
        Name = "SearchLocalCalcEngines";
        ShowIcon = false;
        ShowInTaskbar = false;
        Text = "Search Local CalcEngines...";
        Load += SearchLocalCalcEngines_Load;
        ResumeLayout(false);
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