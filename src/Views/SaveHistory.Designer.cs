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
		this.components = new System.ComponentModel.Container();
		this.label1 = new System.Windows.Forms.Label();
		this.author = new System.Windows.Forms.TextBox();
		this.skip = new System.Windows.Forms.Button();
		this.ok = new System.Windows.Forms.Button();
		this.version = new System.Windows.Forms.TextBox();
		this.versionLabel = new System.Windows.Forms.Label();
		this.description = new System.Windows.Forms.TextBox();
		this.label2 = new System.Windows.Forms.Label();
		this.errorProvider = new System.Windows.Forms.ErrorProvider( this.components );
		this.lManagementSite = new System.Windows.Forms.Label();
		this.tUserName = new System.Windows.Forms.TextBox();
		this.lUserName = new System.Windows.Forms.Label();
		this.tPassword = new System.Windows.Forms.TextBox();
		this.lPassword = new System.Windows.Forms.Label();
		this.forceUpload = new System.Windows.Forms.CheckBox();
		( (System.ComponentModel.ISupportInitialize)( this.errorProvider ) ).BeginInit();
		this.SuspendLayout();
		// 
		// label1
		// 
		this.label1.AutoSize = true;
		this.label1.Location = new System.Drawing.Point( 17, 12 );
		this.label1.Margin = new System.Windows.Forms.Padding( 4, 0, 4, 0 );
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size( 54, 17 );
		this.label1.TabIndex = 0;
		this.label1.Text = "&Author:";
		// 
		// author
		// 
		this.author.Anchor = ( (System.Windows.Forms.AnchorStyles)( ( ( System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left )
		| System.Windows.Forms.AnchorStyles.Right ) ) );
		this.author.Location = new System.Drawing.Point( 21, 32 );
		this.author.Margin = new System.Windows.Forms.Padding( 4 );
		this.author.Name = "author";
		this.author.Size = new System.Drawing.Size( 287, 22 );
		this.author.TabIndex = 1;
		// 
		// skip
		// 
		this.skip.Anchor = ( (System.Windows.Forms.AnchorStyles)( ( System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right ) ) );
		this.skip.DialogResult = System.Windows.Forms.DialogResult.Ignore;
		this.skip.Location = new System.Drawing.Point( 487, 266 );
		this.skip.Margin = new System.Windows.Forms.Padding( 4 );
		this.skip.Name = "skip";
		this.skip.Size = new System.Drawing.Size( 100, 28 );
		this.skip.TabIndex = 13;
		this.skip.Text = "S&kip";
		this.skip.UseVisualStyleBackColor = true;
		// 
		// ok
		// 
		this.ok.Anchor = ( (System.Windows.Forms.AnchorStyles)( ( System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right ) ) );
		this.ok.Location = new System.Drawing.Point( 371, 266 );
		this.ok.Margin = new System.Windows.Forms.Padding( 4 );
		this.ok.Name = "ok";
		this.ok.Size = new System.Drawing.Size( 108, 28 );
		this.ok.TabIndex = 12;
		this.ok.Text = "A&pply/Upload";
		this.ok.UseVisualStyleBackColor = true;
		this.ok.Click += new System.EventHandler( this.Ok_Click );
		// 
		// version
		// 
		this.version.Anchor = ( (System.Windows.Forms.AnchorStyles)( ( System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right ) ) );
		this.version.Location = new System.Drawing.Point( 325, 32 );
		this.version.Margin = new System.Windows.Forms.Padding( 4 );
		this.version.Name = "version";
		this.version.Size = new System.Drawing.Size( 264, 22 );
		this.version.TabIndex = 3;
		// 
		// versionLabel
		// 
		this.versionLabel.Anchor = ( (System.Windows.Forms.AnchorStyles)( ( System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right ) ) );
		this.versionLabel.AutoSize = true;
		this.versionLabel.Location = new System.Drawing.Point( 321, 12 );
		this.versionLabel.Margin = new System.Windows.Forms.Padding( 4, 0, 4, 0 );
		this.versionLabel.Name = "versionLabel";
		this.versionLabel.Size = new System.Drawing.Size( 163, 17 );
		this.versionLabel.TabIndex = 2;
		this.versionLabel.Text = "&Version:";
		// 
		// description
		// 
		this.description.Anchor = ( (System.Windows.Forms.AnchorStyles)( ( ( ( System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom )
		| System.Windows.Forms.AnchorStyles.Left )
		| System.Windows.Forms.AnchorStyles.Right ) ) );
		this.description.Location = new System.Drawing.Point( 21, 82 );
		this.description.Margin = new System.Windows.Forms.Padding( 4 );
		this.description.Multiline = true;
		this.description.Name = "description";
		this.description.Size = new System.Drawing.Size( 566, 62 );
		this.description.TabIndex = 5;
		// 
		// label2
		// 
		this.label2.AutoSize = true;
		this.label2.Location = new System.Drawing.Point( 17, 62 );
		this.label2.Margin = new System.Windows.Forms.Padding( 4, 0, 4, 0 );
		this.label2.Name = "label2";
		this.label2.Size = new System.Drawing.Size( 83, 17 );
		this.label2.TabIndex = 4;
		this.label2.Text = "&Description:";
		// 
		// errorProvider
		// 
		this.errorProvider.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink;
		this.errorProvider.ContainerControl = this;
		// 
		// lManagementSite
		// 
		this.lManagementSite.Anchor = ( (System.Windows.Forms.AnchorStyles)( ( System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left ) ) );
		this.lManagementSite.AutoSize = true;
		this.lManagementSite.Font = new System.Drawing.Font( "Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ( (byte)( 0 ) ) );
		this.lManagementSite.Location = new System.Drawing.Point( 16, 155 );
		this.lManagementSite.Margin = new System.Windows.Forms.Padding( 4, 0, 4, 0 );
		this.lManagementSite.Name = "lManagementSite";
		this.lManagementSite.Size = new System.Drawing.Size( 360, 17 );
		this.lManagementSite.TabIndex = 6;
		this.lManagementSite.Text = "Provide Credentials Upload To Management Site";
		// 
		// tUserName
		// 
		this.tUserName.Anchor = ( (System.Windows.Forms.AnchorStyles)( ( ( System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left )
		| System.Windows.Forms.AnchorStyles.Right ) ) );
		this.tUserName.Location = new System.Drawing.Point( 20, 202 );
		this.tUserName.Margin = new System.Windows.Forms.Padding( 4 );
		this.tUserName.Name = "tUserName";
		this.tUserName.Size = new System.Drawing.Size( 288, 22 );
		this.tUserName.TabIndex = 8;
		// 
		// lUserName
		// 
		this.lUserName.Anchor = ( (System.Windows.Forms.AnchorStyles)( ( System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left ) ) );
		this.lUserName.AutoSize = true;
		this.lUserName.Location = new System.Drawing.Point( 16, 182 );
		this.lUserName.Margin = new System.Windows.Forms.Padding( 4, 0, 4, 0 );
		this.lUserName.Name = "lUserName";
		this.lUserName.Size = new System.Drawing.Size( 83, 17 );
		this.lUserName.TabIndex = 7;
		this.lUserName.Text = "&User Name:";
		// 
		// tPassword
		// 
		this.tPassword.Anchor = ( (System.Windows.Forms.AnchorStyles)( ( System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right ) ) );
		this.tPassword.Location = new System.Drawing.Point( 325, 202 );
		this.tPassword.Margin = new System.Windows.Forms.Padding( 4 );
		this.tPassword.Name = "tPassword";
		this.tPassword.Size = new System.Drawing.Size( 264, 22 );
		this.tPassword.TabIndex = 10;
		this.tPassword.UseSystemPasswordChar = true;
		// 
		// lPassword
		// 
		this.lPassword.Anchor = ( (System.Windows.Forms.AnchorStyles)( ( System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right ) ) );
		this.lPassword.AutoSize = true;
		this.lPassword.Location = new System.Drawing.Point( 321, 182 );
		this.lPassword.Margin = new System.Windows.Forms.Padding( 4, 0, 4, 0 );
		this.lPassword.Name = "lPassword";
		this.lPassword.Size = new System.Drawing.Size( 69, 17 );
		this.lPassword.TabIndex = 9;
		this.lPassword.Text = "&Password";
		// 
		// forceUpload
		// 
		this.forceUpload.Anchor = ( (System.Windows.Forms.AnchorStyles)( ( System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left ) ) );
		this.forceUpload.AutoSize = true;
		this.forceUpload.Location = new System.Drawing.Point( 21, 231 );
		this.forceUpload.Name = "forceUpload";
		this.forceUpload.Size = new System.Drawing.Size( 440, 21 );
		this.forceUpload.TabIndex = 11;
		this.forceUpload.Text = "&Force upload without checking Management Site\'s lastest version";
		this.forceUpload.UseVisualStyleBackColor = true;
		// 
		// SaveHistory
		// 
		this.AcceptButton = this.ok;
		this.AutoScaleDimensions = new System.Drawing.SizeF( 8F, 16F );
		this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.CancelButton = this.skip;
		this.ClientSize = new System.Drawing.Size( 616, 309 );
		this.Controls.Add( this.forceUpload );
		this.Controls.Add( this.tPassword );
		this.Controls.Add( this.lPassword );
		this.Controls.Add( this.tUserName );
		this.Controls.Add( this.lUserName );
		this.Controls.Add( this.lManagementSite );
		this.Controls.Add( this.description );
		this.Controls.Add( this.label2 );
		this.Controls.Add( this.version );
		this.Controls.Add( this.versionLabel );
		this.Controls.Add( this.ok );
		this.Controls.Add( this.skip );
		this.Controls.Add( this.author );
		this.Controls.Add( this.label1 );
		this.Margin = new System.Windows.Forms.Padding( 4 );
		this.MaximizeBox = false;
		this.MaximumSize = new System.Drawing.Size( 927, 481 );
		this.MinimizeBox = false;
		this.MinimumSize = new System.Drawing.Size( 589, 356 );
		this.Name = "SaveHistory";
		this.ShowIcon = false;
		this.ShowInTaskbar = false;
		this.Text = "KAT Version History";
		( (System.ComponentModel.ISupportInitialize)( this.errorProvider ) ).EndInit();
		this.ResumeLayout( false );
		this.PerformLayout();
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