namespace KAT.Camelot.Extensibility.Excel.AddIn;

partial class LoadInputTab
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
        errorProvider = new ErrorProvider(components);
        dataSourceLabel = new Label();
        password = new TextBox();
        passwordLabel = new Label();
        emailAddress = new TextBox();
        emailAddressLabel = new Label();
        ok = new Button();
        cancel = new Button();
        dataSource = new ComboBox();
        inputFileNameSelect = new Button();
        inputFileName = new TextBox();
        inputFileNameLabel = new Label();
        configLookupUrl = new TextBox();
        configLookupsUrlLabel = new Label();
        label3 = new Label();
        utilityTabsDivider = new Label();
        downloadGlobalTables = new CheckBox();
        loadTables = new CheckBox();
        downloadConfigLookups = new CheckBox();
        configLookupPath = new TextBox();
        configLookupPathLabel = new Label();
        configLookupPathSelect = new Button();
        participantSplitContainer = new SplitContainer();
        clientName = new TextBox();
        clientNameLabel = new Label();
        authId = new TextBox();
        authIdLabel = new Label();
        ((System.ComponentModel.ISupportInitialize)errorProvider).BeginInit();
        ((System.ComponentModel.ISupportInitialize)participantSplitContainer).BeginInit();
        participantSplitContainer.Panel1.SuspendLayout();
        participantSplitContainer.Panel2.SuspendLayout();
        participantSplitContainer.SuspendLayout();
        SuspendLayout();
        // 
        // errorProvider
        // 
        errorProvider.BlinkStyle = ErrorBlinkStyle.NeverBlink;
        errorProvider.ContainerControl = this;
        // 
        // dataSourceLabel
        // 
        dataSourceLabel.AutoSize = true;
        dataSourceLabel.Location = new Point(13, 145);
        dataSourceLabel.Margin = new Padding(4, 0, 4, 0);
        dataSourceLabel.Name = "dataSourceLabel";
        dataSourceLabel.Size = new Size(73, 15);
        dataSourceLabel.TabIndex = 5;
        dataSourceLabel.Text = "Data &Source:";
        // 
        // password
        // 
        password.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        password.Location = new Point(13, 95);
        password.Margin = new Padding(4);
        password.Name = "password";
        password.PasswordChar = '*';
        password.Size = new Size(386, 23);
        password.TabIndex = 4;
        // 
        // passwordLabel
        // 
        passwordLabel.AutoSize = true;
        passwordLabel.Location = new Point(13, 76);
        passwordLabel.Margin = new Padding(4, 0, 4, 0);
        passwordLabel.Name = "passwordLabel";
        passwordLabel.Size = new Size(57, 15);
        passwordLabel.TabIndex = 3;
        passwordLabel.Text = "&Password";
        // 
        // emailAddress
        // 
        emailAddress.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        emailAddress.Location = new Point(13, 49);
        emailAddress.Margin = new Padding(4);
        emailAddress.Name = "emailAddress";
        emailAddress.Size = new Size(386, 23);
        emailAddress.TabIndex = 2;
        // 
        // emailAddressLabel
        // 
        emailAddressLabel.AutoSize = true;
        emailAddressLabel.Location = new Point(13, 30);
        emailAddressLabel.Margin = new Padding(4, 0, 4, 0);
        emailAddressLabel.Name = "emailAddressLabel";
        emailAddressLabel.Size = new Size(84, 15);
        emailAddressLabel.TabIndex = 1;
        emailAddressLabel.Text = "&Email Address:";
        // 
        // ok
        // 
        ok.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        ok.Location = new Point(219, 452);
        ok.Margin = new Padding(4);
        ok.Name = "ok";
        ok.Size = new Size(88, 26);
        ok.TabIndex = 19;
        ok.Text = "OK";
        ok.UseVisualStyleBackColor = true;
        ok.Click += Ok_Click;
        // 
        // cancel
        // 
        cancel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        cancel.DialogResult = DialogResult.Cancel;
        cancel.Location = new Point(313, 452);
        cancel.Margin = new Padding(4);
        cancel.Name = "cancel";
        cancel.Size = new Size(88, 26);
        cancel.TabIndex = 20;
        cancel.Text = "Cancel";
        cancel.UseVisualStyleBackColor = true;
        // 
        // dataSource
        // 
        dataSource.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        dataSource.FormattingEnabled = true;
        dataSource.Location = new Point(13, 163);
        dataSource.Name = "dataSource";
        dataSource.Size = new Size(386, 23);
        dataSource.TabIndex = 6;
        dataSource.SelectedIndexChanged += DataSource_SelectedIndexChanged;
		dataSource.DropDownStyle = ComboBoxStyle.DropDownList;
        // 
        // inputFileNameSelect
        // 
        inputFileNameSelect.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        inputFileNameSelect.Location = new Point(361, 207);
        inputFileNameSelect.Margin = new Padding(4, 3, 4, 3);
        inputFileNameSelect.Name = "inputFileNameSelect";
        inputFileNameSelect.Size = new Size(40, 23);
        inputFileNameSelect.TabIndex = 9;
        inputFileNameSelect.Text = "...";
        inputFileNameSelect.UseVisualStyleBackColor = true;
        inputFileNameSelect.Click += InputFileNameSelect_Click;
        // 
        // inputFileName
        // 
        inputFileName.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        inputFileName.Location = new Point(15, 207);
        inputFileName.Margin = new Padding(4, 3, 4, 3);
        inputFileName.Name = "inputFileName";
        inputFileName.Size = new Size(341, 23);
        inputFileName.TabIndex = 8;
        // 
        // inputFileNameLabel
        // 
        inputFileNameLabel.AutoSize = true;
        inputFileNameLabel.Location = new Point(15, 189);
        inputFileNameLabel.Margin = new Padding(4, 0, 4, 0);
        inputFileNameLabel.Name = "inputFileNameLabel";
        inputFileNameLabel.Size = new Size(59, 15);
        inputFileNameLabel.TabIndex = 7;
        inputFileNameLabel.Text = "&Input File:";
        // 
        // configLookupUrl
        // 
        configLookupUrl.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        configLookupUrl.Location = new Point(13, 375);
        configLookupUrl.Margin = new Padding(4);
        configLookupUrl.Name = "configLookupUrl";
        configLookupUrl.Size = new Size(388, 23);
        configLookupUrl.TabIndex = 15;
        // 
        // configLookupsUrlLabel
        // 
        configLookupsUrlLabel.AutoSize = true;
        configLookupsUrlLabel.Location = new Point(13, 357);
        configLookupsUrlLabel.Margin = new Padding(4, 0, 4, 0);
        configLookupsUrlLabel.Name = "configLookupsUrlLabel";
        configLookupsUrlLabel.Size = new Size(171, 15);
        configLookupsUrlLabel.TabIndex = 14;
        configLookupsUrlLabel.Text = "Config-Lookups Download &Url:";
        // 
        // label3
        // 
        label3.AutoSize = true;
        label3.Font = new Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point);
        label3.Location = new Point(13, 9);
        label3.Margin = new Padding(4, 0, 4, 0);
        label3.Name = "label3";
        label3.Size = new Size(156, 13);
        label3.TabIndex = 0;
        label3.Text = "Authentication Information";
        // 
        // utilityTabsDivider
        // 
        utilityTabsDivider.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        utilityTabsDivider.BorderStyle = BorderStyle.Fixed3D;
        utilityTabsDivider.Location = new Point(15, 131);
        utilityTabsDivider.Margin = new Padding(4, 0, 4, 0);
        utilityTabsDivider.Name = "utilityTabsDivider";
        utilityTabsDivider.Size = new Size(386, 2);
        utilityTabsDivider.TabIndex = 21;
        // 
        // downloadGlobalTables
        // 
        downloadGlobalTables.AutoSize = true;
        downloadGlobalTables.Location = new Point(13, 312);
        downloadGlobalTables.Margin = new Padding(4);
        downloadGlobalTables.Name = "downloadGlobalTables";
        downloadGlobalTables.Size = new Size(259, 19);
        downloadGlobalTables.TabIndex = 12;
        downloadGlobalTables.Text = "&Download latest MadHatter_GlobalTables.xls";
        downloadGlobalTables.UseVisualStyleBackColor = true;
        // 
        // loadTables
        // 
        loadTables.AutoSize = true;
        loadTables.Location = new Point(13, 290);
        loadTables.Margin = new Padding(4);
        loadTables.Name = "loadTables";
        loadTables.Size = new Size(224, 19);
        loadTables.TabIndex = 11;
        loadTables.Text = "Load Global and Client &Lookup Tables";
        loadTables.UseVisualStyleBackColor = true;
        loadTables.CheckedChanged += LoadTables_CheckedChanged;
        // 
        // downloadConfigLookups
        // 
        downloadConfigLookups.AutoSize = true;
        downloadConfigLookups.Location = new Point(13, 334);
        downloadConfigLookups.Margin = new Padding(4);
        downloadConfigLookups.Name = "downloadConfigLookups";
        downloadConfigLookups.Size = new Size(192, 19);
        downloadConfigLookups.TabIndex = 13;
        downloadConfigLookups.Text = "&Download Config-Lookups.xml";
        downloadConfigLookups.UseVisualStyleBackColor = true;
        downloadConfigLookups.CheckedChanged += DownloadConfigLookups_CheckedChanged;
        // 
        // configLookupPath
        // 
        configLookupPath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        configLookupPath.Location = new Point(15, 420);
        configLookupPath.Margin = new Padding(4);
        configLookupPath.Name = "configLookupPath";
        configLookupPath.Size = new Size(341, 23);
        configLookupPath.TabIndex = 17;
        // 
        // configLookupPathLabel
        // 
        configLookupPathLabel.AutoSize = true;
        configLookupPathLabel.Location = new Point(15, 402);
        configLookupPathLabel.Margin = new Padding(4, 0, 4, 0);
        configLookupPathLabel.Name = "configLookupPathLabel";
        configLookupPathLabel.Size = new Size(154, 15);
        configLookupPathLabel.TabIndex = 16;
        configLookupPathLabel.Text = "Local Config-Lookups &Path:";
        // 
        // configLookupPathSelect
        // 
        configLookupPathSelect.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        configLookupPathSelect.Location = new Point(361, 420);
        configLookupPathSelect.Margin = new Padding(4, 3, 4, 3);
        configLookupPathSelect.Name = "configLookupPathSelect";
        configLookupPathSelect.Size = new Size(40, 23);
        configLookupPathSelect.TabIndex = 18;
        configLookupPathSelect.Text = "...";
        configLookupPathSelect.UseVisualStyleBackColor = true;
        configLookupPathSelect.Click += ConfigLookupPathSelect_Click;
        // 
        // participantSplitContainer
        // 
        participantSplitContainer.Anchor = AnchorStyles.Left | AnchorStyles.Right;
        participantSplitContainer.IsSplitterFixed = true;
        participantSplitContainer.Location = new Point(13, 236);
        participantSplitContainer.Name = "participantSplitContainer";
        // 
        // participantSplitContainer.Panel1
        // 
        participantSplitContainer.Panel1.Controls.Add(clientName);
        participantSplitContainer.Panel1.Controls.Add(clientNameLabel);
        // 
        // participantSplitContainer.Panel2
        // 
        participantSplitContainer.Panel2.Controls.Add(authId);
        participantSplitContainer.Panel2.Controls.Add(authIdLabel);
        participantSplitContainer.Size = new Size(386, 51);
        participantSplitContainer.SplitterDistance = 191;
        participantSplitContainer.TabIndex = 10;
        participantSplitContainer.Resize += ParticipantSplitContainer_Resize;
        // 
        // clientName
        // 
        clientName.Anchor = AnchorStyles.Left | AnchorStyles.Right;
        clientName.Location = new Point(2, 23);
        clientName.Margin = new Padding(4);
        clientName.Name = "clientName";
        clientName.Size = new Size(190, 23);
        clientName.TabIndex = 1;
        // 
        // clientNameLabel
        // 
        clientNameLabel.AutoSize = true;
        clientNameLabel.Location = new Point(0, 4);
        clientNameLabel.Margin = new Padding(4, 0, 4, 0);
        clientNameLabel.Name = "clientNameLabel";
        clientNameLabel.Size = new Size(73, 15);
        clientNameLabel.TabIndex = 0;
        clientNameLabel.Text = "&Client Name";
        // 
        // authId
        // 
        authId.Anchor = AnchorStyles.Left | AnchorStyles.Right;
        authId.Location = new Point(4, 23);
        authId.Margin = new Padding(4);
        authId.Name = "authId";
        authId.Size = new Size(189, 23);
        authId.TabIndex = 1;
        // 
        // authIdLabel
        // 
        authIdLabel.AutoSize = true;
        authIdLabel.Location = new Point(4, 4);
        authIdLabel.Margin = new Padding(4, 0, 4, 0);
        authIdLabel.Name = "authIdLabel";
        authIdLabel.Size = new Size(50, 15);
        authIdLabel.TabIndex = 0;
        authIdLabel.Text = "Auth &ID:";
        // 
        // LoadInputTab
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(414, 491);
        Controls.Add(participantSplitContainer);
        Controls.Add(configLookupPathSelect);
        Controls.Add(configLookupPath);
        Controls.Add(configLookupPathLabel);
        Controls.Add(downloadConfigLookups);
        Controls.Add(loadTables);
        Controls.Add(downloadGlobalTables);
        Controls.Add(utilityTabsDivider);
        Controls.Add(label3);
        Controls.Add(configLookupUrl);
        Controls.Add(configLookupsUrlLabel);
        Controls.Add(inputFileNameSelect);
        Controls.Add(inputFileName);
        Controls.Add(inputFileNameLabel);
        Controls.Add(dataSource);
        Controls.Add(dataSourceLabel);
        Controls.Add(password);
        Controls.Add(passwordLabel);
        Controls.Add(emailAddress);
        Controls.Add(emailAddressLabel);
        Controls.Add(ok);
        Controls.Add(cancel);
        Margin = new Padding(4);
        MaximizeBox = false;
        MaximumSize = new Size(510, 530);
        MinimizeBox = false;
        MinimumSize = new Size(309, 530);
        Name = "LoadInputTab";
        ShowIcon = false;
        ShowInTaskbar = false;
        SizeGripStyle = SizeGripStyle.Show;
        Text = "Load Input Data";
        Load += LoadInputTab_Load;
        ((System.ComponentModel.ISupportInitialize)errorProvider).EndInit();
        participantSplitContainer.Panel1.ResumeLayout(false);
        participantSplitContainer.Panel1.PerformLayout();
        participantSplitContainer.Panel2.ResumeLayout(false);
        participantSplitContainer.Panel2.PerformLayout();
        ((System.ComponentModel.ISupportInitialize)participantSplitContainer).EndInit();
        participantSplitContainer.ResumeLayout(false);
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion
    private System.Windows.Forms.ErrorProvider errorProvider;
	private System.Windows.Forms.Label dataSourceLabel;
	private System.Windows.Forms.Label passwordLabel;
	private System.Windows.Forms.TextBox password;
	private System.Windows.Forms.Label emailAddressLabel;
	private System.Windows.Forms.TextBox emailAddress;
	private System.Windows.Forms.Button ok;
	private System.Windows.Forms.Button cancel;
    private ComboBox dataSource;
    private TextBox configLookupUrl;
    private Label configLookupsUrlLabel;
    private Button inputFileNameSelect;
    private TextBox inputFileName;
    private Label inputFileNameLabel;
    private Label label3;
    private Label utilityTabsDivider;
    private CheckBox downloadGlobalTables;
    private CheckBox loadTables;
    private CheckBox downloadConfigLookups;
    private TextBox configLookupPath;
    private Label configLookupPathLabel;
    private Button configLookupPathSelect;
    private SplitContainer participantSplitContainer;
    private TextBox clientName;
    private Label clientNameLabel;
    private TextBox authId;
    private Label authIdLabel;
}