
namespace ArduinoSerialLogger
{
    partial class Form1
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
            this.requestDataButton = new System.Windows.Forms.Button();
            this.portBox = new System.Windows.Forms.ListBox();
            this.baudrateNumericUpDown = new System.Windows.Forms.NumericUpDown();
            this.refreshPortsButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.requestContinuousDataButton = new System.Windows.Forms.Button();
            this.saveDataCheckBox = new System.Windows.Forms.CheckBox();
            this.closeConnectionButton = new System.Windows.Forms.Button();
            this.saveLocationLabel = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.delimiterComboBox = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.writeToExcelCheckbox = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.nextColumnRadioButton = new System.Windows.Forms.RadioButton();
            this.nextRowRadioButton = new System.Windows.Forms.RadioButton();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.lineDelimiterComboBox = new System.Windows.Forms.ComboBox();
            this.moveActiveCellCheckBox = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.baudrateNumericUpDown)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.SuspendLayout();
            // 
            // requestDataButton
            // 
            this.requestDataButton.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.requestDataButton.Location = new System.Drawing.Point(6, 55);
            this.requestDataButton.Name = "requestDataButton";
            this.requestDataButton.Size = new System.Drawing.Size(200, 30);
            this.requestDataButton.TabIndex = 0;
            this.requestDataButton.Text = "Request Data";
            this.requestDataButton.UseVisualStyleBackColor = true;
            this.requestDataButton.Click += new System.EventHandler(this.requestDataButton_Click);
            // 
            // portBox
            // 
            this.portBox.FormattingEnabled = true;
            this.portBox.ItemHeight = 16;
            this.portBox.Location = new System.Drawing.Point(6, 50);
            this.portBox.Name = "portBox";
            this.portBox.Size = new System.Drawing.Size(120, 68);
            this.portBox.TabIndex = 1;
            this.portBox.SelectedIndexChanged += new System.EventHandler(this.portBox_SelectedIndexChanged);
            // 
            // baudrateNumericUpDown
            // 
            this.baudrateNumericUpDown.Location = new System.Drawing.Point(135, 45);
            this.baudrateNumericUpDown.Maximum = new decimal(new int[] {
            100000000,
            0,
            0,
            0});
            this.baudrateNumericUpDown.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.baudrateNumericUpDown.Name = "baudrateNumericUpDown";
            this.baudrateNumericUpDown.Size = new System.Drawing.Size(120, 22);
            this.baudrateNumericUpDown.TabIndex = 3;
            this.baudrateNumericUpDown.Value = new decimal(new int[] {
            9600,
            0,
            0,
            0});
            this.baudrateNumericUpDown.ValueChanged += new System.EventHandler(this.baudrateNumericUpDown_ValueChanged);
            // 
            // refreshPortsButton
            // 
            this.refreshPortsButton.Location = new System.Drawing.Point(6, 21);
            this.refreshPortsButton.Name = "refreshPortsButton";
            this.refreshPortsButton.Size = new System.Drawing.Size(120, 25);
            this.refreshPortsButton.TabIndex = 4;
            this.refreshPortsButton.Text = "Refresh Ports";
            this.refreshPortsButton.UseVisualStyleBackColor = true;
            this.refreshPortsButton.Click += new System.EventHandler(this.refreshPortsButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(132, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 17);
            this.label1.TabIndex = 5;
            this.label1.Text = "Baudrate";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // requestContinuousDataButton
            // 
            this.requestContinuousDataButton.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.requestContinuousDataButton.Location = new System.Drawing.Point(6, 22);
            this.requestContinuousDataButton.Name = "requestContinuousDataButton";
            this.requestContinuousDataButton.Size = new System.Drawing.Size(200, 30);
            this.requestContinuousDataButton.TabIndex = 6;
            this.requestContinuousDataButton.Text = "Log Continuous Data";
            this.requestContinuousDataButton.UseVisualStyleBackColor = true;
            this.requestContinuousDataButton.Click += new System.EventHandler(this.requestContinuousDataButton_Click);
            // 
            // saveDataCheckBox
            // 
            this.saveDataCheckBox.AutoSize = true;
            this.saveDataCheckBox.Location = new System.Drawing.Point(6, 27);
            this.saveDataCheckBox.Name = "saveDataCheckBox";
            this.saveDataCheckBox.Size = new System.Drawing.Size(143, 21);
            this.saveDataCheckBox.TabIndex = 7;
            this.saveDataCheckBox.Text = "Save Data To File";
            this.saveDataCheckBox.UseVisualStyleBackColor = true;
            this.saveDataCheckBox.CheckedChanged += new System.EventHandler(this.saveDataCheckBox_CheckedChanged);
            // 
            // closeConnectionButton
            // 
            this.closeConnectionButton.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.closeConnectionButton.Enabled = false;
            this.closeConnectionButton.Location = new System.Drawing.Point(6, 88);
            this.closeConnectionButton.Name = "closeConnectionButton";
            this.closeConnectionButton.Size = new System.Drawing.Size(200, 30);
            this.closeConnectionButton.TabIndex = 8;
            this.closeConnectionButton.Text = "Close Connection";
            this.closeConnectionButton.UseVisualStyleBackColor = true;
            this.closeConnectionButton.Click += new System.EventHandler(this.closeConnectionButton_Click);
            // 
            // saveLocationLabel
            // 
            this.saveLocationLabel.AutoSize = true;
            this.saveLocationLabel.Location = new System.Drawing.Point(6, 88);
            this.saveLocationLabel.Name = "saveLocationLabel";
            this.saveLocationLabel.Size = new System.Drawing.Size(91, 17);
            this.saveLocationLabel.TabIndex = 9;
            this.saveLocationLabel.Text = "save location";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(6, 54);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(150, 25);
            this.button1.TabIndex = 10;
            this.button1.Text = "Save Location";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.saveLocationLabel_Click);
            // 
            // delimiterComboBox
            // 
            this.delimiterComboBox.FormattingEnabled = true;
            this.delimiterComboBox.Location = new System.Drawing.Point(135, 94);
            this.delimiterComboBox.Name = "delimiterComboBox";
            this.delimiterComboBox.Size = new System.Drawing.Size(120, 24);
            this.delimiterComboBox.TabIndex = 11;
            this.delimiterComboBox.SelectedIndexChanged += new System.EventHandler(this.delimiterComboBox_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(132, 74);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(116, 17);
            this.label2.TabIndex = 12;
            this.label2.Text = "Dataset Delimiter";
            // 
            // writeToExcelCheckbox
            // 
            this.writeToExcelCheckbox.AutoSize = true;
            this.writeToExcelCheckbox.Location = new System.Drawing.Point(6, 21);
            this.writeToExcelCheckbox.Name = "writeToExcelCheckbox";
            this.writeToExcelCheckbox.Size = new System.Drawing.Size(185, 21);
            this.writeToExcelCheckbox.TabIndex = 13;
            this.writeToExcelCheckbox.Text = "Write to Active Excel Cell";
            this.writeToExcelCheckbox.UseVisualStyleBackColor = true;
            this.writeToExcelCheckbox.CheckedChanged += new System.EventHandler(this.writeToExcelCheckbox_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.moveActiveCellCheckBox);
            this.groupBox1.Controls.Add(this.nextColumnRadioButton);
            this.groupBox1.Controls.Add(this.nextRowRadioButton);
            this.groupBox1.Location = new System.Drawing.Point(6, 50);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(255, 112);
            this.groupBox1.TabIndex = 16;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Move Active Cell after Dataset?";
            // 
            // nextColumnRadioButton
            // 
            this.nextColumnRadioButton.AutoSize = true;
            this.nextColumnRadioButton.Location = new System.Drawing.Point(6, 74);
            this.nextColumnRadioButton.Name = "nextColumnRadioButton";
            this.nextColumnRadioButton.Size = new System.Drawing.Size(111, 21);
            this.nextColumnRadioButton.TabIndex = 17;
            this.nextColumnRadioButton.Text = "Next Collumn";
            this.nextColumnRadioButton.UseVisualStyleBackColor = true;
            this.nextColumnRadioButton.CheckedChanged += new System.EventHandler(this.nextColumnRadioButton_CheckedChanged);
            // 
            // nextRowRadioButton
            // 
            this.nextRowRadioButton.AutoSize = true;
            this.nextRowRadioButton.Checked = true;
            this.nextRowRadioButton.Location = new System.Drawing.Point(6, 47);
            this.nextRowRadioButton.Name = "nextRowRadioButton";
            this.nextRowRadioButton.Size = new System.Drawing.Size(88, 21);
            this.nextRowRadioButton.TabIndex = 16;
            this.nextRowRadioButton.TabStop = true;
            this.nextRowRadioButton.Text = "Next Row";
            this.nextRowRadioButton.UseVisualStyleBackColor = true;
            this.nextRowRadioButton.CheckedChanged += new System.EventHandler(this.nextRowRadioButton_CheckedChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.refreshPortsButton);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.portBox);
            this.groupBox2.Controls.Add(this.delimiterComboBox);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.baudrateNumericUpDown);
            this.groupBox2.Location = new System.Drawing.Point(12, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(276, 129);
            this.groupBox2.TabIndex = 17;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "General Settings";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.groupBox6);
            this.groupBox3.Controls.Add(this.writeToExcelCheckbox);
            this.groupBox3.Controls.Add(this.groupBox1);
            this.groupBox3.Location = new System.Drawing.Point(18, 156);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(270, 232);
            this.groupBox3.TabIndex = 18;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Excel Settings";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.button1);
            this.groupBox4.Controls.Add(this.saveLocationLabel);
            this.groupBox4.Controls.Add(this.saveDataCheckBox);
            this.groupBox4.Location = new System.Drawing.Point(294, 21);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(258, 120);
            this.groupBox4.TabIndex = 19;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Logger Settings";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.closeConnectionButton);
            this.groupBox5.Controls.Add(this.requestDataButton);
            this.groupBox5.Controls.Add(this.requestContinuousDataButton);
            this.groupBox5.Location = new System.Drawing.Point(294, 156);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(258, 162);
            this.groupBox5.TabIndex = 20;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Logger Control";
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.lineDelimiterComboBox);
            this.groupBox6.Location = new System.Drawing.Point(6, 168);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(255, 52);
            this.groupBox6.TabIndex = 18;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Datapoint Delimiter";
            // 
            // lineDelimiterComboBox
            // 
            this.lineDelimiterComboBox.FormattingEnabled = true;
            this.lineDelimiterComboBox.Location = new System.Drawing.Point(6, 21);
            this.lineDelimiterComboBox.Name = "lineDelimiterComboBox";
            this.lineDelimiterComboBox.Size = new System.Drawing.Size(121, 24);
            this.lineDelimiterComboBox.TabIndex = 0;
            this.lineDelimiterComboBox.SelectedIndexChanged += new System.EventHandler(this.lineDelimiterComboBox_SelectedIndexChanged_1);
            // 
            // moveActiveCellCheckBox
            // 
            this.moveActiveCellCheckBox.AutoSize = true;
            this.moveActiveCellCheckBox.Location = new System.Drawing.Point(6, 21);
            this.moveActiveCellCheckBox.Name = "moveActiveCellCheckBox";
            this.moveActiveCellCheckBox.Size = new System.Drawing.Size(133, 21);
            this.moveActiveCellCheckBox.TabIndex = 18;
            this.moveActiveCellCheckBox.Text = "Move Active Cell";
            this.moveActiveCellCheckBox.UseVisualStyleBackColor = true;
            this.moveActiveCellCheckBox.CheckedChanged += new System.EventHandler(this.moveActiveCellCheckBox_CheckedChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Name = "Form1";
            this.Text = "Form1";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.baudrateNumericUpDown)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox6.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button requestDataButton;
        private System.Windows.Forms.ListBox portBox;
        private System.Windows.Forms.NumericUpDown baudrateNumericUpDown;
        private System.Windows.Forms.Button refreshPortsButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button requestContinuousDataButton;
        private System.Windows.Forms.CheckBox saveDataCheckBox;
        private System.Windows.Forms.Button closeConnectionButton;
        private System.Windows.Forms.Label saveLocationLabel;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox delimiterComboBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox writeToExcelCheckbox;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton nextColumnRadioButton;
        private System.Windows.Forms.RadioButton nextRowRadioButton;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.ComboBox lineDelimiterComboBox;
        private System.Windows.Forms.CheckBox moveActiveCellCheckBox;
    }
}

