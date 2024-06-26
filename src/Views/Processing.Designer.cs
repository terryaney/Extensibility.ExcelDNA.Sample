namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal partial class Processing
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
        cancel = new Button();
        processingLabel = new Label();
        progressBar = new ProgressBar();
        SuspendLayout();
        // 
        // cancel
        // 
        cancel.DialogResult = DialogResult.Ignore;
        cancel.Location = new Point(235, 81);
        cancel.Margin = new Padding(4);
        cancel.Name = "cancel";
        cancel.Size = new Size(88, 26);
        cancel.TabIndex = 2;
        cancel.Text = "Cancel";
        cancel.UseVisualStyleBackColor = true;
		cancel.Click += Cancel_Click;
        // 
        // processingLabel
        // 
        processingLabel.AutoSize = true;
        processingLabel.Location = new Point(11, 23);
        processingLabel.Margin = new Padding(4, 0, 4, 0);
        processingLabel.Name = "processingLabel";
        processingLabel.Size = new Size(73, 15);
        processingLabel.TabIndex = 0;
        processingLabel.Text = "Processing...";
        // 
        // progressBar
        // 
        progressBar.Location = new Point(12, 51);
        progressBar.Name = "progressBar";
        progressBar.Size = new Size(311, 23);
        progressBar.TabIndex = 1;
        // 
        // Processing
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        CancelButton = cancel;
        ClientSize = new Size(336, 136);
        Controls.Add(progressBar);
        Controls.Add(processingLabel);
        Controls.Add(cancel);
        Margin = new Padding(4);
        MaximizeBox = false;
        MaximumSize = new Size(352, 175);
        MinimizeBox = false;
        MinimumSize = new Size(352, 175);
        Name = "Processing";
        ShowIcon = false;
        ShowInTaskbar = false;
        SizeGripStyle = SizeGripStyle.Hide;
        Text = "Progress";
        Load += Processing_Load;
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion
    private System.Windows.Forms.Button cancel;
	private System.Windows.Forms.Label processingLabel;
    private ProgressBar progressBar;
}