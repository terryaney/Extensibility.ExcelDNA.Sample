namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal partial class Credentials
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
        errorProvider = new ErrorProvider(components);
        tUserName = new TextBox();
        lUserName = new Label();
        tPassword = new TextBox();
        lPassword = new Label();
        ((System.ComponentModel.ISupportInitialize)errorProvider).BeginInit();
        SuspendLayout();
        // 
        // cancel
        // 
        cancel.DialogResult = DialogResult.Ignore;
        cancel.Location = new Point(235, 97);
        cancel.Margin = new Padding(4, 4, 4, 4);
        cancel.Name = "cancel";
        cancel.Size = new Size(88, 26);
        cancel.TabIndex = 5;
        cancel.Text = "Cancel";
        cancel.UseVisualStyleBackColor = true;
        // 
        // ok
        // 
        ok.Location = new Point(134, 97);
        ok.Margin = new Padding(4, 4, 4, 4);
        ok.Name = "ok";
        ok.Size = new Size(94, 26);
        ok.TabIndex = 4;
        ok.Text = "OK";
        ok.UseVisualStyleBackColor = true;
        ok.Click += Ok_Click;
        // 
        // errorProvider
        // 
        errorProvider.BlinkStyle = ErrorBlinkStyle.NeverBlink;
        errorProvider.ContainerControl = this;
        // 
        // tUserName
        // 
        tUserName.Location = new Point(15, 25);
        tUserName.Margin = new Padding(4, 4, 4, 4);
        tUserName.Name = "tUserName";
        tUserName.Size = new Size(308, 23);
        tUserName.TabIndex = 1;
        // 
        // lUserName
        // 
        lUserName.AutoSize = true;
        lUserName.Location = new Point(11, 7);
        lUserName.Margin = new Padding(4, 0, 4, 0);
        lUserName.Name = "lUserName";
        lUserName.Size = new Size(68, 15);
        lUserName.TabIndex = 0;
        lUserName.Text = "&User Name:";
        // 
        // tPassword
        // 
        tPassword.Location = new Point(15, 69);
        tPassword.Margin = new Padding(4, 4, 4, 4);
        tPassword.Name = "tPassword";
        tPassword.Size = new Size(308, 23);
        tPassword.TabIndex = 3;
        tPassword.UseSystemPasswordChar = true;
        // 
        // lPassword
        // 
        lPassword.AutoSize = true;
        lPassword.Location = new Point(11, 51);
        lPassword.Margin = new Padding(4, 0, 4, 0);
        lPassword.Name = "lPassword";
        lPassword.Size = new Size(57, 15);
        lPassword.TabIndex = 2;
        lPassword.Text = "&Password";
        // 
        // Credentials
        // 
        AcceptButton = ok;
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        CancelButton = cancel;
        ClientSize = new Size(336, 136);
        Controls.Add(tPassword);
        Controls.Add(lPassword);
        Controls.Add(tUserName);
        Controls.Add(lUserName);
        Controls.Add(ok);
        Controls.Add(cancel);
        Margin = new Padding(4, 4, 4, 4);
        MaximizeBox = false;
        MaximumSize = new Size(352, 175);
        MinimizeBox = false;
        MinimumSize = new Size(352, 175);
        Name = "Credentials";
        ShowIcon = false;
        ShowInTaskbar = false;
        SizeGripStyle = SizeGripStyle.Hide;
        Text = "KAT Credentials";
        Load += Credentials_Load;
        ((System.ComponentModel.ISupportInitialize)errorProvider).EndInit();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion
    private System.Windows.Forms.Button cancel;
	private System.Windows.Forms.Button ok;
	private System.Windows.Forms.ErrorProvider errorProvider;
	private System.Windows.Forms.TextBox tPassword;
	private System.Windows.Forms.Label lPassword;
	private System.Windows.Forms.TextBox tUserName;
	private System.Windows.Forms.Label lUserName;
}