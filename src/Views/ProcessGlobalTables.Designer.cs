namespace KAT.Camelot.Extensibility.Excel.AddIn;

partial class ProcessGlobalTables
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
        targetsLabel = new Label();
        password = new TextBox();
        passwordLabel = new Label();
        emailAddress = new TextBox();
        emailAddressLabel = new Label();
        ok = new Button();
        cancel = new Button();
        clientName = new TextBox();
        clientNameLabel = new Label();
        targets = new CheckedListBox();
        ((System.ComponentModel.ISupportInitialize)errorProvider).BeginInit();
        SuspendLayout();
        // 
        // errorProvider
        // 
        errorProvider.BlinkStyle = ErrorBlinkStyle.NeverBlink;
        errorProvider.ContainerControl = this;
        // 
        // targetsLabel
        // 
        targetsLabel.AutoSize = true;
        targetsLabel.Location = new Point(13, 9);
        targetsLabel.Margin = new Padding(4, 0, 4, 0);
        targetsLabel.Name = "targetsLabel";
        targetsLabel.Size = new Size(44, 15);
        targetsLabel.TabIndex = 0;
        targetsLabel.Text = "&Targets";
        // 
        // password
        // 
        password.Location = new Point(15, 252);
        password.Margin = new Padding(4);
        password.Name = "password";
        password.PasswordChar = '*';
        password.Size = new Size(237, 23);
        password.TabIndex = 7;
        // 
        // passwordLabel
        // 
        passwordLabel.AutoSize = true;
        passwordLabel.Location = new Point(15, 233);
        passwordLabel.Margin = new Padding(4, 0, 4, 0);
        passwordLabel.Name = "passwordLabel";
        passwordLabel.Size = new Size(57, 15);
        passwordLabel.TabIndex = 6;
        passwordLabel.Text = "&Password";
        // 
        // emailAddress
        // 
        emailAddress.Location = new Point(15, 206);
        emailAddress.Margin = new Padding(4);
        emailAddress.Name = "emailAddress";
        emailAddress.Size = new Size(237, 23);
        emailAddress.TabIndex = 5;
        // 
        // emailAddressLabel
        // 
        emailAddressLabel.AutoSize = true;
        emailAddressLabel.Location = new Point(15, 187);
        emailAddressLabel.Margin = new Padding(4, 0, 4, 0);
        emailAddressLabel.Name = "emailAddressLabel";
        emailAddressLabel.Size = new Size(84, 15);
        emailAddressLabel.TabIndex = 4;
        emailAddressLabel.Text = "&Email Address:";
        // 
        // ok
        // 
        ok.Location = new Point(70, 283);
        ok.Margin = new Padding(4);
        ok.Name = "ok";
        ok.Size = new Size(88, 26);
        ok.TabIndex = 8;
        ok.Text = "OK";
        ok.UseVisualStyleBackColor = true;
        ok.Click += Ok_Click;
        // 
        // cancel
        // 
        cancel.DialogResult = DialogResult.Cancel;
        cancel.Location = new Point(164, 283);
        cancel.Margin = new Padding(4);
        cancel.Name = "cancel";
        cancel.Size = new Size(88, 26);
        cancel.TabIndex = 9;
        cancel.Text = "Cancel";
        cancel.UseVisualStyleBackColor = true;
        // 
        // clientName
        // 
        clientName.Location = new Point(13, 160);
        clientName.Margin = new Padding(4);
        clientName.Name = "clientName";
        clientName.Size = new Size(239, 23);
        clientName.TabIndex = 3;
        // 
        // clientNameLabel
        // 
        clientNameLabel.AutoSize = true;
        clientNameLabel.Location = new Point(13, 142);
        clientNameLabel.Margin = new Padding(4, 0, 4, 0);
        clientNameLabel.Name = "clientNameLabel";
        clientNameLabel.Size = new Size(73, 15);
        clientNameLabel.TabIndex = 2;
        clientNameLabel.Text = "&Client Name";
        // 
        // targets
        // 
        targets.CheckOnClick = true;
        targets.FormattingEnabled = true;
        targets.Location = new Point(13, 27);
        targets.Name = "targets";
        targets.Size = new Size(237, 112);
        targets.TabIndex = 1;
        // 
        // ProcessGlobalTables
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(274, 322);
        Controls.Add(targets);
        Controls.Add(targetsLabel);
        Controls.Add(password);
        Controls.Add(passwordLabel);
        Controls.Add(emailAddress);
        Controls.Add(emailAddressLabel);
        Controls.Add(ok);
        Controls.Add(cancel);
        Controls.Add(clientName);
        Controls.Add(clientNameLabel);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        Margin = new Padding(4);
        MaximizeBox = false;
        MaximumSize = new Size(813, 453);
        MinimizeBox = false;
        MinimumSize = new Size(100, 100);
        Name = "ProcessGlobalTables";
        ShowIcon = false;
        ShowInTaskbar = false;
        SizeGripStyle = SizeGripStyle.Show;
        Text = "Process Global Tables";
        Load += ProcessGlobalTable_Load;
        ((System.ComponentModel.ISupportInitialize)errorProvider).EndInit();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion
    private System.Windows.Forms.ErrorProvider errorProvider;
	private System.Windows.Forms.Label targetsLabel;
	private CheckedListBox targets;
	private System.Windows.Forms.Label passwordLabel;
	private System.Windows.Forms.TextBox password;
	private System.Windows.Forms.Label emailAddressLabel;
	private System.Windows.Forms.TextBox emailAddress;
	private System.Windows.Forms.Label clientNameLabel;
	private System.Windows.Forms.TextBox clientName;
	private System.Windows.Forms.Button ok;
	private System.Windows.Forms.Button cancel;
}