
namespace ExcelVerifier
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.inputLabel = new System.Windows.Forms.Label();
            this.outputLabel = new System.Windows.Forms.Label();
            this.inputTextBox = new System.Windows.Forms.TextBox();
            this.outputTextBox = new System.Windows.Forms.TextBox();
            this.validateButton = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.outputFormatLabel = new System.Windows.Forms.Label();
            this.inputFormatLabel = new System.Windows.Forms.Label();
            this.outputFormat = new System.Windows.Forms.ComboBox();
            this.inputFormat = new System.Windows.Forms.ComboBox();
            this.exportButton = new System.Windows.Forms.Button();
            this.outputBrowseButton = new System.Windows.Forms.Button();
            this.inputBrowseButton = new System.Windows.Forms.Button();
            this.resultLabel = new System.Windows.Forms.Label();
            this.statusLabel = new System.Windows.Forms.Label();
            this.ddmmyyRadioButton = new System.Windows.Forms.RadioButton();
            this.mmddyyRadioButton = new System.Windows.Forms.RadioButton();
            this.dateFormatLabel = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // inputLabel
            // 
            this.inputLabel.AutoSize = true;
            this.inputLabel.Location = new System.Drawing.Point(15, 36);
            this.inputLabel.Name = "inputLabel";
            this.inputLabel.Size = new System.Drawing.Size(104, 13);
            this.inputLabel.TabIndex = 0;
            this.inputLabel.Text = "Input Excel Location";
            // 
            // outputLabel
            // 
            this.outputLabel.AutoSize = true;
            this.outputLabel.Location = new System.Drawing.Point(15, 81);
            this.outputLabel.Name = "outputLabel";
            this.outputLabel.Size = new System.Drawing.Size(112, 13);
            this.outputLabel.TabIndex = 1;
            this.outputLabel.Text = "Output Excel Location";
            // 
            // inputTextBox
            // 
            this.inputTextBox.Location = new System.Drawing.Point(147, 33);
            this.inputTextBox.Name = "inputTextBox";
            this.inputTextBox.Size = new System.Drawing.Size(261, 20);
            this.inputTextBox.TabIndex = 2;
            // 
            // outputTextBox
            // 
            this.outputTextBox.Location = new System.Drawing.Point(147, 78);
            this.outputTextBox.Name = "outputTextBox";
            this.outputTextBox.Size = new System.Drawing.Size(261, 20);
            this.outputTextBox.TabIndex = 3;
            // 
            // validateButton
            // 
            this.validateButton.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.validateButton.Location = new System.Drawing.Point(211, 407);
            this.validateButton.Name = "validateButton";
            this.validateButton.Size = new System.Drawing.Size(122, 33);
            this.validateButton.TabIndex = 4;
            this.validateButton.Text = "Validate";
            this.validateButton.UseVisualStyleBackColor = false;
            this.validateButton.Click += new System.EventHandler(this.validateButton_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dateFormatLabel);
            this.groupBox1.Controls.Add(this.mmddyyRadioButton);
            this.groupBox1.Controls.Add(this.ddmmyyRadioButton);
            this.groupBox1.Controls.Add(this.outputFormatLabel);
            this.groupBox1.Controls.Add(this.inputFormatLabel);
            this.groupBox1.Controls.Add(this.outputFormat);
            this.groupBox1.Controls.Add(this.inputFormat);
            this.groupBox1.Controls.Add(this.exportButton);
            this.groupBox1.Controls.Add(this.outputBrowseButton);
            this.groupBox1.Controls.Add(this.inputBrowseButton);
            this.groupBox1.Controls.Add(this.inputLabel);
            this.groupBox1.Controls.Add(this.validateButton);
            this.groupBox1.Controls.Add(this.outputLabel);
            this.groupBox1.Controls.Add(this.outputTextBox);
            this.groupBox1.Controls.Add(this.inputTextBox);
            this.groupBox1.Location = new System.Drawing.Point(41, 42);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(505, 494);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Validation";
            // 
            // outputFormatLabel
            // 
            this.outputFormatLabel.AutoSize = true;
            this.outputFormatLabel.Location = new System.Drawing.Point(15, 214);
            this.outputFormatLabel.Name = "outputFormatLabel";
            this.outputFormatLabel.Size = new System.Drawing.Size(103, 13);
            this.outputFormatLabel.TabIndex = 12;
            this.outputFormatLabel.Text = "Output Excel Format";
            // 
            // inputFormatLabel
            // 
            this.inputFormatLabel.AutoSize = true;
            this.inputFormatLabel.Location = new System.Drawing.Point(15, 159);
            this.inputFormatLabel.Name = "inputFormatLabel";
            this.inputFormatLabel.Size = new System.Drawing.Size(95, 13);
            this.inputFormatLabel.TabIndex = 11;
            this.inputFormatLabel.Text = "Input Excel Format";
            // 
            // outputFormat
            // 
            this.outputFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.outputFormat.FormattingEnabled = true;
            this.outputFormat.Items.AddRange(new object[] {
            "VIA MEDICA",
            "LDM"});
            this.outputFormat.Location = new System.Drawing.Point(147, 211);
            this.outputFormat.Name = "outputFormat";
            this.outputFormat.Size = new System.Drawing.Size(156, 21);
            this.outputFormat.TabIndex = 10;
            // 
            // inputFormat
            // 
            this.inputFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.inputFormat.FormattingEnabled = true;
            this.inputFormat.Items.AddRange(new object[] {
            "VIA MEDICA"});
            this.inputFormat.Location = new System.Drawing.Point(147, 156);
            this.inputFormat.Name = "inputFormat";
            this.inputFormat.Size = new System.Drawing.Size(156, 21);
            this.inputFormat.TabIndex = 9;
            // 
            // exportButton
            // 
            this.exportButton.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.exportButton.Location = new System.Drawing.Point(351, 407);
            this.exportButton.Name = "exportButton";
            this.exportButton.Size = new System.Drawing.Size(122, 33);
            this.exportButton.TabIndex = 8;
            this.exportButton.Text = "Export";
            this.exportButton.UseVisualStyleBackColor = false;
            this.exportButton.Click += new System.EventHandler(this.exportButton_Click);
            // 
            // outputBrowseButton
            // 
            this.outputBrowseButton.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.outputBrowseButton.Location = new System.Drawing.Point(414, 76);
            this.outputBrowseButton.Name = "outputBrowseButton";
            this.outputBrowseButton.Size = new System.Drawing.Size(52, 23);
            this.outputBrowseButton.TabIndex = 7;
            this.outputBrowseButton.Text = "Browse";
            this.outputBrowseButton.UseVisualStyleBackColor = false;
            this.outputBrowseButton.Click += new System.EventHandler(this.outputBrowseButton_Click);
            // 
            // inputBrowseButton
            // 
            this.inputBrowseButton.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.inputBrowseButton.Location = new System.Drawing.Point(414, 31);
            this.inputBrowseButton.Name = "inputBrowseButton";
            this.inputBrowseButton.Size = new System.Drawing.Size(52, 23);
            this.inputBrowseButton.TabIndex = 6;
            this.inputBrowseButton.Text = "Browse";
            this.inputBrowseButton.UseVisualStyleBackColor = false;
            this.inputBrowseButton.Click += new System.EventHandler(this.inputBrowseButton_Click);
            // 
            // resultLabel
            // 
            this.resultLabel.AutoSize = true;
            this.resultLabel.Location = new System.Drawing.Point(487, 17);
            this.resultLabel.Name = "resultLabel";
            this.resultLabel.Size = new System.Drawing.Size(0, 13);
            this.resultLabel.TabIndex = 6;
            // 
            // statusLabel
            // 
            this.statusLabel.AutoSize = true;
            this.statusLabel.Font = new System.Drawing.Font("Unispace", 9.749999F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.statusLabel.ForeColor = System.Drawing.Color.Yellow;
            this.statusLabel.Location = new System.Drawing.Point(56, 425);
            this.statusLabel.MaximumSize = new System.Drawing.Size(510, 0);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(0, 15);
            this.statusLabel.TabIndex = 9;
            // 
            // ddmmyyRadioButton
            // 
            this.ddmmyyRadioButton.AutoSize = true;
            this.ddmmyyRadioButton.Checked = true;
            this.ddmmyyRadioButton.Location = new System.Drawing.Point(147, 298);
            this.ddmmyyRadioButton.Name = "ddmmyyRadioButton";
            this.ddmmyyRadioButton.Size = new System.Drawing.Size(83, 17);
            this.ddmmyyRadioButton.TabIndex = 13;
            this.ddmmyyRadioButton.TabStop = true;
            this.ddmmyyRadioButton.Text = "dd/mm/yyyy";
            this.ddmmyyRadioButton.UseVisualStyleBackColor = true;
            // 
            // mmddyyRadioButton
            // 
            this.mmddyyRadioButton.AutoSize = true;
            this.mmddyyRadioButton.Location = new System.Drawing.Point(259, 298);
            this.mmddyyRadioButton.Name = "mmddyyRadioButton";
            this.mmddyyRadioButton.Size = new System.Drawing.Size(83, 17);
            this.mmddyyRadioButton.TabIndex = 14;
            this.mmddyyRadioButton.Text = "mm/dd/yyyy";
            this.mmddyyRadioButton.UseVisualStyleBackColor = true;
            // 
            // dateFormatLabel
            // 
            this.dateFormatLabel.AutoSize = true;
            this.dateFormatLabel.Location = new System.Drawing.Point(15, 295);
            this.dateFormatLabel.MaximumSize = new System.Drawing.Size(100, 0);
            this.dateFormatLabel.Name = "dateFormatLabel";
            this.dateFormatLabel.Size = new System.Drawing.Size(90, 26);
            this.dateFormatLabel.TabIndex = 15;
            this.dateFormatLabel.Text = "Initial date format (in input file)";
            this.dateFormatLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.ClientSize = new System.Drawing.Size(590, 697);
            this.Controls.Add(this.statusLabel);
            this.Controls.Add(this.resultLabel);
            this.Controls.Add(this.groupBox1);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel Validator v1.0";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label inputLabel;
        private System.Windows.Forms.Label outputLabel;
        private System.Windows.Forms.TextBox inputTextBox;
        private System.Windows.Forms.TextBox outputTextBox;
        private System.Windows.Forms.Button validateButton;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label resultLabel;
        private System.Windows.Forms.Button outputBrowseButton;
        private System.Windows.Forms.Button inputBrowseButton;
        private System.Windows.Forms.Button exportButton;
        private System.Windows.Forms.Label statusLabel;
        private System.Windows.Forms.ComboBox outputFormat;
        private System.Windows.Forms.ComboBox inputFormat;
        private System.Windows.Forms.Label outputFormatLabel;
        private System.Windows.Forms.Label inputFormatLabel;
        private System.Windows.Forms.Label dateFormatLabel;
        private System.Windows.Forms.RadioButton mmddyyRadioButton;
        private System.Windows.Forms.RadioButton ddmmyyRadioButton;
    }
}

