namespace KAT.Camelot.Extensibility.Excel.AddIn;

partial class XmlMapping
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
		inputFileName = new TextBox();
		inputFileNameSelect = new Button();
		cancel = new Button();
		ok = new Button();
		outputFileNameSelect = new Button();
		outputFileName = new TextBox();
		label3 = new Label();
		errorProvider = new ErrorProvider( components );
		clientName = new TextBox();
		label2 = new Label();
		( (System.ComponentModel.ISupportInitialize)errorProvider ).BeginInit();
		SuspendLayout();
		// 
		// label1
		// 
		label1.AutoSize = true;
		label1.Location = new Point( 14, 60 );
		label1.Margin = new Padding( 4, 0, 4, 0 );
		label1.Name = "label1";
		label1.Size = new Size( 59, 15 );
		label1.TabIndex = 2;
		label1.Text = "&Input File:";
		// 
		// inputFileName
		// 
		inputFileName.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		inputFileName.Location = new Point( 14, 78 );
		inputFileName.Margin = new Padding( 4, 3, 4, 3 );
		inputFileName.Name = "inputFileName";
		inputFileName.Size = new Size( 421, 23 );
		inputFileName.TabIndex = 3;
		// 
		// inputFileNameSelect
		// 
		inputFileNameSelect.Anchor = AnchorStyles.Top | AnchorStyles.Right;
		inputFileNameSelect.Location = new Point( 443, 78 );
		inputFileNameSelect.Margin = new Padding( 4, 3, 4, 3 );
		inputFileNameSelect.Name = "inputFileNameSelect";
		inputFileNameSelect.Size = new Size( 40, 23 );
		inputFileNameSelect.TabIndex = 4;
		inputFileNameSelect.Text = "...";
		inputFileNameSelect.UseVisualStyleBackColor = true;
		inputFileNameSelect.Click += InputFileNameSelect_Click;
		// 
		// cancel
		// 
		cancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		cancel.DialogResult = DialogResult.Cancel;
		cancel.Location = new Point( 395, 170 );
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
		ok.Location = new Point( 301, 170 );
		ok.Margin = new Padding( 4, 3, 4, 3 );
		ok.Name = "ok";
		ok.Size = new Size( 88, 27 );
		ok.TabIndex = 8;
		ok.Text = "OK";
		ok.UseVisualStyleBackColor = true;
		ok.Click += Ok_Click;
		// 
		// outputFileNameSelect
		// 
		outputFileNameSelect.Anchor = AnchorStyles.Top | AnchorStyles.Right;
		outputFileNameSelect.Location = new Point( 443, 128 );
		outputFileNameSelect.Margin = new Padding( 4, 3, 4, 3 );
		outputFileNameSelect.Name = "outputFileNameSelect";
		outputFileNameSelect.Size = new Size( 40, 23 );
		outputFileNameSelect.TabIndex = 7;
		outputFileNameSelect.Text = "...";
		outputFileNameSelect.UseVisualStyleBackColor = true;
		outputFileNameSelect.Click += OuputFileNameSelect_Click;
		// 
		// outputFileName
		// 
		outputFileName.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		outputFileName.Location = new Point( 14, 128 );
		outputFileName.Margin = new Padding( 4, 3, 4, 3 );
		outputFileName.Name = "outputFileName";
		outputFileName.Size = new Size( 421, 23 );
		outputFileName.TabIndex = 6;
		// 
		// label3
		// 
		label3.AutoSize = true;
		label3.Location = new Point( 14, 110 );
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
		clientName.Location = new Point( 14, 29 );
		clientName.Margin = new Padding( 4, 3, 4, 3 );
		clientName.Name = "clientName";
		clientName.Size = new Size( 469, 23 );
		clientName.TabIndex = 1;
		// 
		// label2
		// 
		label2.AutoSize = true;
		label2.Location = new Point( 14, 11 );
		label2.Margin = new Padding( 4, 0, 4, 0 );
		label2.Name = "label2";
		label2.Size = new Size( 73, 15 );
		label2.TabIndex = 0;
		label2.Text = "&Client Name";
		// 
		// XmlMapping
		// 
		AcceptButton = ok;
		AutoScaleDimensions = new SizeF( 7F, 15F );
		AutoScaleMode = AutoScaleMode.Font;
		CancelButton = cancel;
		ClientSize = new Size( 500, 209 );
		Controls.Add( clientName );
		Controls.Add( label2 );
		Controls.Add( outputFileNameSelect );
		Controls.Add( outputFileName );
		Controls.Add( label3 );
		Controls.Add( ok );
		Controls.Add( cancel );
		Controls.Add( inputFileNameSelect );
		Controls.Add( inputFileName );
		Controls.Add( label1 );
		Margin = new Padding( 4, 3, 4, 3 );
		MaximizeBox = false;
		MaximumSize = new Size( 814, 248 );
		MinimizeBox = false;
		MinimumSize = new Size( 371, 248 );
		Name = "XmlMapping";
		ShowIcon = false;
		ShowInTaskbar = false;
		Text = "Xml Data Mapping";
		Load += XmlMapping_Load;
		( (System.ComponentModel.ISupportInitialize)errorProvider ).EndInit();
		ResumeLayout( false );
		PerformLayout();
	}

	#endregion

	private System.Windows.Forms.Label label1;
	private System.Windows.Forms.TextBox inputFileName;
	private System.Windows.Forms.Button inputFileNameSelect;
	private System.Windows.Forms.Button cancel;
	private System.Windows.Forms.Button ok;
	private System.Windows.Forms.Button outputFileNameSelect;
	private System.Windows.Forms.TextBox outputFileName;
	private System.Windows.Forms.Label label3;
	private System.Windows.Forms.ErrorProvider errorProvider;
	private System.Windows.Forms.TextBox clientName;
	private System.Windows.Forms.Label label2;
}