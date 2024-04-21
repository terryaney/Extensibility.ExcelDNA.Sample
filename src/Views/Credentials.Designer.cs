namespace KAT.Extensibility.Excel.AddIn;

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
		errorProvider = new ErrorProvider( components );
		tUserName = new TextBox();
		lUserName = new Label();
		tPassword = new TextBox();
		lPassword = new Label();
		( (System.ComponentModel.ISupportInitialize)errorProvider ).BeginInit();
		SuspendLayout();
		// 
		// cancel
		// 
		cancel.DialogResult = DialogResult.Ignore;
		cancel.Location = new Point( 269, 129 );
		cancel.Margin = new Padding( 4, 5, 4, 5 );
		cancel.Name = "cancel";
		cancel.Size = new Size( 100, 35 );
		cancel.TabIndex = 13;
		cancel.Text = "Cancel";
		cancel.UseVisualStyleBackColor = true;
		// 
		// ok
		// 
		ok.Location = new Point( 153, 129 );
		ok.Margin = new Padding( 4, 5, 4, 5 );
		ok.Name = "ok";
		ok.Size = new Size( 108, 35 );
		ok.TabIndex = 12;
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
		tUserName.Location = new Point( 17, 33 );
		tUserName.Margin = new Padding( 4, 5, 4, 5 );
		tUserName.Name = "tUserName";
		tUserName.Size = new Size( 352, 27 );
		tUserName.TabIndex = 8;
		// 
		// lUserName
		// 
		lUserName.AutoSize = true;
		lUserName.Location = new Point( 13, 9 );
		lUserName.Margin = new Padding( 4, 0, 4, 0 );
		lUserName.Name = "lUserName";
		lUserName.Size = new Size( 85, 20 );
		lUserName.TabIndex = 7;
		lUserName.Text = "&User Name:";
		// 
		// tPassword
		// 
		tPassword.Location = new Point( 17, 92 );
		tPassword.Margin = new Padding( 4, 5, 4, 5 );
		tPassword.Name = "tPassword";
		tPassword.Size = new Size( 352, 27 );
		tPassword.TabIndex = 10;
		tPassword.UseSystemPasswordChar = true;
		// 
		// lPassword
		// 
		lPassword.AutoSize = true;
		lPassword.Location = new Point( 13, 68 );
		lPassword.Margin = new Padding( 4, 0, 4, 0 );
		lPassword.Name = "lPassword";
		lPassword.Size = new Size( 70, 20 );
		lPassword.TabIndex = 9;
		lPassword.Text = "&Password";
		// 
		// Credentials
		// 
		AcceptButton = ok;
		AutoScaleDimensions = new SizeF( 8F, 20F );
		AutoScaleMode = AutoScaleMode.Font;
		CancelButton = cancel;
		ClientSize = new Size( 382, 174 );
		Controls.Add( tPassword );
		Controls.Add( lPassword );
		Controls.Add( tUserName );
		Controls.Add( lUserName );
		Controls.Add( ok );
		Controls.Add( cancel );
		Margin = new Padding( 4, 5, 4, 5 );
		MaximizeBox = false;
		MaximumSize = new Size( 400, 221 );
		MinimizeBox = false;
		MinimumSize = new Size( 400, 221 );
		Name = "Credentials";
		ShowIcon = false;
		ShowInTaskbar = false;
		SizeGripStyle = SizeGripStyle.Hide;
		Text = "KAT Credentials";
		( (System.ComponentModel.ISupportInitialize)errorProvider ).EndInit();
		ResumeLayout( false );
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