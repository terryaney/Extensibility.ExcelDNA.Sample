namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal partial class SaveHistory
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
        label1 = new Label();
        author = new TextBox();
        skip = new Button();
        ok = new Button();
        version = new TextBox();
        versionLabel = new Label();
        description = new TextBox();
        label2 = new Label();
        errorProvider = new ErrorProvider(components);
        lManagementSite = new Label();
        tUserName = new TextBox();
        lUserName = new Label();
        tPassword = new TextBox();
        lPassword = new Label();
        forceUpload = new CheckBox();
        ((System.ComponentModel.ISupportInitialize)errorProvider).BeginInit();
        SuspendLayout();
        // 
        // label1
        // 
        label1.AutoSize = true;
        label1.Location = new Point(15, 11);
        label1.Margin = new Padding(4, 0, 4, 0);
        label1.Name = "label1";
        label1.Size = new Size(47, 15);
        label1.TabIndex = 0;
        label1.Text = "&Author:";
        // 
        // author
        // 
        author.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        author.Location = new Point(18, 30);
        author.Margin = new Padding(4);
        author.Name = "author";
        author.Size = new Size(252, 23);
        author.TabIndex = 1;
        // 
        // skip
        // 
        skip.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        skip.DialogResult = DialogResult.Ignore;
        skip.Location = new Point(426, 249);
        skip.Margin = new Padding(4);
        skip.Name = "skip";
        skip.Size = new Size(88, 26);
        skip.TabIndex = 13;
        skip.Text = "S&kip";
        skip.UseVisualStyleBackColor = true;
        // 
        // ok
        // 
        ok.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        ok.Location = new Point(325, 249);
        ok.Margin = new Padding(4);
        ok.Name = "ok";
        ok.Size = new Size(94, 26);
        ok.TabIndex = 12;
        ok.Text = "A&pply/Upload";
        ok.UseVisualStyleBackColor = true;
        ok.Click += Ok_Click;
        // 
        // version
        // 
        version.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        version.Location = new Point(284, 30);
        version.Margin = new Padding(4);
        version.Name = "version";
        version.Size = new Size(232, 23);
        version.TabIndex = 3;
        // 
        // versionLabel
        // 
        versionLabel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        versionLabel.AutoSize = true;
        versionLabel.Location = new Point(281, 11);
        versionLabel.Margin = new Padding(4, 0, 4, 0);
        versionLabel.Name = "versionLabel";
        versionLabel.Size = new Size(48, 15);
        versionLabel.TabIndex = 2;
        versionLabel.Text = "&Version:";
        // 
        // description
        // 
        description.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        description.Location = new Point(18, 77);
        description.Margin = new Padding(4);
        description.Multiline = true;
        description.Name = "description";
        description.Size = new Size(496, 58);
        description.TabIndex = 5;
        // 
        // label2
        // 
        label2.AutoSize = true;
        label2.Location = new Point(15, 58);
        label2.Margin = new Padding(4, 0, 4, 0);
        label2.Name = "label2";
        label2.Size = new Size(70, 15);
        label2.TabIndex = 4;
        label2.Text = "&Description:";
        // 
        // errorProvider
        // 
        errorProvider.BlinkStyle = ErrorBlinkStyle.NeverBlink;
        errorProvider.ContainerControl = this;
        // 
        // lManagementSite
        // 
        lManagementSite.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        lManagementSite.AutoSize = true;
        lManagementSite.Font = new Font("Microsoft Sans Serif", 7.8F, FontStyle.Bold, GraphicsUnit.Point);
        lManagementSite.Location = new Point(14, 145);
        lManagementSite.Margin = new Padding(4, 0, 4, 0);
        lManagementSite.Name = "lManagementSite";
        lManagementSite.Size = new Size(282, 13);
        lManagementSite.TabIndex = 6;
        lManagementSite.Text = "Provide Credentials Upload To Management Site";
        // 
        // tUserName
        // 
        tUserName.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        tUserName.Location = new Point(18, 189);
        tUserName.Margin = new Padding(4);
        tUserName.Name = "tUserName";
        tUserName.Size = new Size(252, 23);
        tUserName.TabIndex = 8;
        // 
        // lUserName
        // 
        lUserName.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        lUserName.AutoSize = true;
        lUserName.Location = new Point(14, 171);
        lUserName.Margin = new Padding(4, 0, 4, 0);
        lUserName.Name = "lUserName";
        lUserName.Size = new Size(68, 15);
        lUserName.TabIndex = 7;
        lUserName.Text = "&User Name:";
        // 
        // tPassword
        // 
        tPassword.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        tPassword.Location = new Point(284, 189);
        tPassword.Margin = new Padding(4);
        tPassword.Name = "tPassword";
        tPassword.Size = new Size(232, 23);
        tPassword.TabIndex = 10;
        tPassword.UseSystemPasswordChar = true;
        // 
        // lPassword
        // 
        lPassword.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        lPassword.AutoSize = true;
        lPassword.Location = new Point(281, 171);
        lPassword.Margin = new Padding(4, 0, 4, 0);
        lPassword.Name = "lPassword";
        lPassword.Size = new Size(57, 15);
        lPassword.TabIndex = 9;
        lPassword.Text = "&Password";
        // 
        // forceUpload
        // 
        forceUpload.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        forceUpload.AutoSize = true;
        forceUpload.Location = new Point(18, 218);
        forceUpload.Name = "forceUpload";
        forceUpload.Size = new Size(371, 19);
        forceUpload.TabIndex = 11;
        forceUpload.Text = "&Force upload without checking Management Site's lastest version";
        forceUpload.UseVisualStyleBackColor = true;
        // 
        // SaveHistory
        // 
        AcceptButton = ok;
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        CancelButton = skip;
        ClientSize = new Size(539, 297);
        Controls.Add(forceUpload);
        Controls.Add(tPassword);
        Controls.Add(lPassword);
        Controls.Add(tUserName);
        Controls.Add(lUserName);
        Controls.Add(lManagementSite);
        Controls.Add(description);
        Controls.Add(label2);
        Controls.Add(version);
        Controls.Add(versionLabel);
        Controls.Add(ok);
        Controls.Add(skip);
        Controls.Add(author);
        Controls.Add(label1);
        Margin = new Padding(4);
        MaximizeBox = false;
        MaximumSize = new Size(813, 453);
        MinimizeBox = false;
        MinimumSize = new Size(517, 336);
        Name = "SaveHistory";
        ShowIcon = false;
        ShowInTaskbar = false;
        Text = "KAT Version History";
        Load += SaveHistory_Load;
        ((System.ComponentModel.ISupportInitialize)errorProvider).EndInit();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private System.Windows.Forms.Label label1;
	private System.Windows.Forms.TextBox author;
	private System.Windows.Forms.Button skip;
	private System.Windows.Forms.Button ok;
	private System.Windows.Forms.TextBox version;
	private System.Windows.Forms.Label versionLabel;
	private System.Windows.Forms.TextBox description;
	private System.Windows.Forms.Label label2;
	private System.Windows.Forms.ErrorProvider errorProvider;
	private System.Windows.Forms.TextBox tPassword;
	private System.Windows.Forms.Label lPassword;
	private System.Windows.Forms.TextBox tUserName;
	private System.Windows.Forms.Label lUserName;
	private System.Windows.Forms.Label lManagementSite;
	private System.Windows.Forms.CheckBox forceUpload;
}