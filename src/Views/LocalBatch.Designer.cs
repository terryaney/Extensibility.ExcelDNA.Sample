namespace KAT.Camelot.Extensibility.Excel.AddIn;

partial class LocalBatch
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
        errorProvider = new ErrorProvider(components);
        filter = new TextBox();
        filterLabel = new Label();
        inputTab = new ComboBox();
        inputTabLabel = new Label();
        resultTab = new ComboBox();
        resultTabLabel = new Label();
        exportType = new ComboBox();
        exportTypeLabel = new Label();
        limitRowsTo = new TextBox();
        limitRows = new CheckBox();
        saveErrorCalcEngineError = new CheckBox();
        saveErrorCalcEngineCount = new TextBox();
        ((System.ComponentModel.ISupportInitialize)errorProvider).BeginInit();
        SuspendLayout();
        // 
        // label1
        // 
        label1.AutoSize = true;
        label1.Location = new Point(13, 9);
        label1.Margin = new Padding(4, 0, 4, 0);
        label1.Name = "label1";
        label1.Size = new Size(59, 15);
        label1.TabIndex = 0;
        label1.Text = "&Input File:";
        // 
        // inputFileName
        // 
        inputFileName.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        inputFileName.Location = new Point(13, 27);
        inputFileName.Margin = new Padding(4, 3, 4, 3);
        inputFileName.Name = "inputFileName";
        inputFileName.Size = new Size(341, 23);
        inputFileName.TabIndex = 1;
        // 
        // inputFileNameSelect
        // 
        inputFileNameSelect.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        inputFileNameSelect.Location = new Point(362, 27);
        inputFileNameSelect.Margin = new Padding(4, 3, 4, 3);
        inputFileNameSelect.Name = "inputFileNameSelect";
        inputFileNameSelect.Size = new Size(40, 23);
        inputFileNameSelect.TabIndex = 2;
        inputFileNameSelect.Text = "...";
        inputFileNameSelect.UseVisualStyleBackColor = true;
        inputFileNameSelect.Click += InputFileNameSelect_Click;
        // 
        // cancel
        // 
        cancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        cancel.DialogResult = DialogResult.Cancel;
        cancel.Location = new Point(319, 351);
        cancel.Margin = new Padding(4, 3, 4, 3);
        cancel.Name = "cancel";
        cancel.Size = new Size(88, 27);
        cancel.TabIndex = 19;
        cancel.Text = "Cancel";
        cancel.UseVisualStyleBackColor = true;
        // 
        // ok
        // 
        ok.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        ok.Location = new Point(225, 351);
        ok.Margin = new Padding(4, 3, 4, 3);
        ok.Name = "ok";
        ok.Size = new Size(88, 27);
        ok.TabIndex = 18;
        ok.Text = "OK";
        ok.UseVisualStyleBackColor = true;
        ok.Click += Ok_Click;
        // 
        // outputFileNameSelect
        // 
        outputFileNameSelect.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        outputFileNameSelect.Location = new Point(362, 247);
        outputFileNameSelect.Margin = new Padding(4, 3, 4, 3);
        outputFileNameSelect.Name = "outputFileNameSelect";
        outputFileNameSelect.Size = new Size(40, 23);
        outputFileNameSelect.TabIndex = 13;
        outputFileNameSelect.Text = "...";
        outputFileNameSelect.UseVisualStyleBackColor = true;
        outputFileNameSelect.Click += OuputFileNameSelect_Click;
        // 
        // outputFileName
        // 
        outputFileName.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        outputFileName.Location = new Point(12, 247);
        outputFileName.Margin = new Padding(4, 3, 4, 3);
        outputFileName.Name = "outputFileName";
        outputFileName.Size = new Size(341, 23);
        outputFileName.TabIndex = 12;
        // 
        // label3
        // 
        label3.AutoSize = true;
        label3.Location = new Point(12, 229);
        label3.Margin = new Padding(4, 0, 4, 0);
        label3.Name = "label3";
        label3.Size = new Size(69, 15);
        label3.TabIndex = 11;
        label3.Text = "&Output File:";
        // 
        // errorProvider
        // 
        errorProvider.BlinkStyle = ErrorBlinkStyle.NeverBlink;
        errorProvider.ContainerControl = this;
        // 
        // filter
        // 
        filter.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        filter.Location = new Point(13, 71);
        filter.Margin = new Padding(4, 3, 4, 3);
        filter.Name = "filter";
        filter.Size = new Size(389, 23);
        filter.TabIndex = 4;
		filter.PlaceholderText = "HistoryData/HistoryItem[@hisType='Status'][position()=last()]/status='A'";
		// 
		// filterLabel
		// 
		filterLabel.AutoSize = true;
        filterLabel.Location = new Point(13, 53);
        filterLabel.Margin = new Padding(4, 0, 4, 0);
        filterLabel.Name = "filterLabel";
        filterLabel.Size = new Size(36, 15);
        filterLabel.TabIndex = 3;
        filterLabel.Text = "&Filter:";
        // 
        // inputTab
        // 
        inputTab.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        inputTab.DropDownStyle = ComboBoxStyle.DropDownList;
        inputTab.FormattingEnabled = true;
        inputTab.Location = new Point(15, 115);
        inputTab.Name = "inputTab";
        inputTab.Size = new Size(387, 23);
        inputTab.TabIndex = 6;
        // 
        // inputTabLabel
        // 
        inputTabLabel.AutoSize = true;
        inputTabLabel.Location = new Point(13, 97);
        inputTabLabel.Margin = new Padding(4, 0, 4, 0);
        inputTabLabel.Name = "inputTabLabel";
        inputTabLabel.Size = new Size(59, 15);
        inputTabLabel.TabIndex = 5;
        inputTabLabel.Text = "Input &Tab:";
        // 
        // resultTab
        // 
        resultTab.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        resultTab.DropDownStyle = ComboBoxStyle.DropDownList;
        resultTab.FormattingEnabled = true;
        resultTab.Location = new Point(13, 159);
        resultTab.Name = "resultTab";
        resultTab.Size = new Size(387, 23);
        resultTab.TabIndex = 8;
        // 
        // resultTabLabel
        // 
        resultTabLabel.AutoSize = true;
        resultTabLabel.Location = new Point(13, 141);
        resultTabLabel.Margin = new Padding(4, 0, 4, 0);
        resultTabLabel.Name = "resultTabLabel";
        resultTabLabel.Size = new Size(63, 15);
        resultTabLabel.TabIndex = 7;
        resultTabLabel.Text = "&Result Tab:";
        // 
        // exportType
        // 
        exportType.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        exportType.DropDownStyle = ComboBoxStyle.DropDownList;
        exportType.FormattingEnabled = true;
        exportType.Location = new Point(12, 203);
        exportType.Name = "exportType";
        exportType.Size = new Size(387, 23);
        exportType.TabIndex = 10;
        // 
        // exportTypeLabel
        // 
        exportTypeLabel.AutoSize = true;
        exportTypeLabel.Location = new Point(12, 185);
        exportTypeLabel.Margin = new Padding(4, 0, 4, 0);
        exportTypeLabel.Name = "exportTypeLabel";
        exportTypeLabel.Size = new Size(71, 15);
        exportTypeLabel.TabIndex = 9;
        exportTypeLabel.Text = "Export &Type:";
        // 
        // limitRowsTo
        // 
        limitRowsTo.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        limitRowsTo.Location = new Point(288, 281);
        limitRowsTo.Margin = new Padding(4, 3, 4, 3);
        limitRowsTo.Name = "limitRowsTo";
        limitRowsTo.Size = new Size(111, 23);
        limitRowsTo.TabIndex = 15;
        // 
        // limitRows
        // 
        limitRows.AutoSize = true;
        limitRows.Location = new Point(13, 283);
        limitRows.Margin = new Padding(4);
        limitRows.Name = "limitRows";
        limitRows.Size = new Size(254, 19);
        limitRows.TabIndex = 14;
        limitRows.Text = "&Limit to xDS Input Data to Number of Rows";
        limitRows.UseVisualStyleBackColor = true;
        limitRows.CheckedChanged += LimitRows_CheckedChanged;
        // 
        // saveErrorCalcEngineError
        // 
        saveErrorCalcEngineError.AutoSize = true;
        saveErrorCalcEngineError.Location = new Point(12, 312);
        saveErrorCalcEngineError.Margin = new Padding(4);
        saveErrorCalcEngineError.Name = "saveErrorCalcEngineError";
        saveErrorCalcEngineError.Size = new Size(220, 19);
        saveErrorCalcEngineError.TabIndex = 16;
        saveErrorCalcEngineError.Text = "&Number of Error CalcEngines to Save";
        saveErrorCalcEngineError.UseVisualStyleBackColor = true;
        saveErrorCalcEngineError.CheckedChanged += SaveErrorCalcEngineError_CheckedChanged;
        // 
        // saveErrorCalcEngineCount
        // 
        saveErrorCalcEngineCount.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        saveErrorCalcEngineCount.Location = new Point(289, 310);
        saveErrorCalcEngineCount.Margin = new Padding(4, 3, 4, 3);
        saveErrorCalcEngineCount.Name = "saveErrorCalcEngineCount";
        saveErrorCalcEngineCount.Size = new Size(111, 23);
        saveErrorCalcEngineCount.TabIndex = 17;
        // 
        // LocalBatch
        // 
        AcceptButton = ok;
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        CancelButton = cancel;
        ClientSize = new Size(420, 390);
        Controls.Add(saveErrorCalcEngineCount);
        Controls.Add(saveErrorCalcEngineError);
        Controls.Add(limitRowsTo);
        Controls.Add(limitRows);
        Controls.Add(exportType);
        Controls.Add(exportTypeLabel);
        Controls.Add(resultTab);
        Controls.Add(resultTabLabel);
        Controls.Add(inputTab);
        Controls.Add(inputTabLabel);
        Controls.Add(filter);
        Controls.Add(filterLabel);
        Controls.Add(outputFileNameSelect);
        Controls.Add(outputFileName);
        Controls.Add(label3);
        Controls.Add(ok);
        Controls.Add(cancel);
        Controls.Add(inputFileNameSelect);
        Controls.Add(inputFileName);
        Controls.Add(label1);
        Margin = new Padding(4, 3, 4, 3);
        MaximizeBox = false;
        MaximumSize = new Size(601, 429);
        MinimizeBox = false;
        MinimumSize = new Size(436, 429);
        Name = "LocalBatch";
        ShowIcon = false;
        ShowInTaskbar = false;
        Text = "Local Batch Calculation";
        Load += LocalBatch_Load;
        ((System.ComponentModel.ISupportInitialize)errorProvider).EndInit();
        ResumeLayout(false);
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
	private System.Windows.Forms.TextBox filter;
	private System.Windows.Forms.Label filterLabel;
    private ComboBox resultTab;
    private Label resultTabLabel;
    private ComboBox inputTab;
    private Label inputTabLabel;
    private ComboBox exportType;
    private Label exportTypeLabel;
    private CheckBox saveErrorCalcEngineError;
    private TextBox limitRowsTo;
    private CheckBox limitRows;
    private TextBox saveErrorCalcEngineCount;
}