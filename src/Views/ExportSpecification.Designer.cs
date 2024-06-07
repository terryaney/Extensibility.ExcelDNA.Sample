namespace KAT.Camelot.Extensibility.Excel.AddIn;

partial class ExportSpecification
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
        locationsLabel = new Label();
        ok = new Button();
        cancel = new Button();
        locations = new CheckedListBox();
        saveSpecification = new CheckBox();
        addLocation = new Button();
        ((System.ComponentModel.ISupportInitialize)errorProvider).BeginInit();
        SuspendLayout();
        // 
        // errorProvider
        // 
        errorProvider.BlinkStyle = ErrorBlinkStyle.NeverBlink;
        errorProvider.ContainerControl = this;
        // 
        // locationsLabel
        // 
        locationsLabel.AutoSize = true;
        locationsLabel.Location = new Point(13, 9);
        locationsLabel.Margin = new Padding(4, 0, 4, 0);
        locationsLabel.Name = "locationsLabel";
        locationsLabel.Size = new Size(138, 15);
        locationsLabel.TabIndex = 0;
        locationsLabel.Text = "&Configuration Locations:";
        // 
        // ok
        // 
        ok.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        ok.Location = new Point(416, 293);
        ok.Margin = new Padding(4);
        ok.Name = "ok";
        ok.Size = new Size(88, 26);
        ok.TabIndex = 4;
        ok.Text = "OK";
        ok.UseVisualStyleBackColor = true;
        ok.Click += Ok_Click;
        // 
        // cancel
        // 
        cancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        cancel.DialogResult = DialogResult.Cancel;
        cancel.Location = new Point(510, 293);
        cancel.Margin = new Padding(4);
        cancel.Name = "cancel";
        cancel.Size = new Size(88, 26);
        cancel.TabIndex = 5;
        cancel.Text = "Cancel";
        cancel.UseVisualStyleBackColor = true;
        // 
        // locations
        // 
        locations.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        locations.CheckOnClick = true;
        locations.FormattingEnabled = true;
        locations.Location = new Point(13, 27);
        locations.Name = "locations";
        locations.Size = new Size(583, 220);
        locations.TabIndex = 1;
        locations.KeyDown += Locations_KeyDown;
        // 
        // saveSpecification
        // 
        saveSpecification.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        saveSpecification.AutoSize = true;
        saveSpecification.Location = new Point(13, 254);
        saveSpecification.Margin = new Padding(4);
        saveSpecification.Name = "saveSpecification";
        saveSpecification.Size = new Size(438, 19);
        saveSpecification.TabIndex = 2;
        saveSpecification.Text = "&Save Specification to \\Evolution\\Websites\\Admin\\HarrisSeverance\\_Developer";
        saveSpecification.UseVisualStyleBackColor = true;
        // 
        // addLocation
        // 
        addLocation.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        addLocation.Location = new Point(13, 293);
        addLocation.Margin = new Padding(4);
        addLocation.Name = "addLocation";
        addLocation.Size = new Size(127, 26);
        addLocation.TabIndex = 3;
        addLocation.Text = "&Add Location...";
        addLocation.UseVisualStyleBackColor = true;
        addLocation.Click += AddLocation_Click;
        // 
        // ExportSpecification
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        CancelButton = cancel;
        ClientSize = new Size(620, 332);
        Controls.Add(addLocation);
        Controls.Add(saveSpecification);
        Controls.Add(locations);
        Controls.Add(locationsLabel);
        Controls.Add(ok);
        Controls.Add(cancel);
        Margin = new Padding(4);
        MaximizeBox = false;
        MaximumSize = new Size(813, 453);
        MinimizeBox = false;
        MinimumSize = new Size(636, 371);
        Name = "ExportSpecification";
        ShowIcon = false;
        ShowInTaskbar = false;
        SizeGripStyle = SizeGripStyle.Show;
        Text = "Export Specification Configuration";
        Load += ExportSpecification_Load;
        ((System.ComponentModel.ISupportInitialize)errorProvider).EndInit();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion
    private System.Windows.Forms.ErrorProvider errorProvider;
	private System.Windows.Forms.Label locationsLabel;
	private CheckedListBox locations;
	private System.Windows.Forms.Button ok;
	private System.Windows.Forms.Button cancel;
    private CheckBox saveSpecification;
    private Button addLocation;
}