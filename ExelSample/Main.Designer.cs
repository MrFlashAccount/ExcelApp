namespace ExelSample
{
    partial class Main
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
            this.OpenFullReportButton = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.FullReportPathTextBox = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.InOutReportPathTextBox = new System.Windows.Forms.TextBox();
            this.OpenInOutReportButton = new System.Windows.Forms.Button();
            this.FireButton = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.CheckSchedule = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // OpenFullReportButton
            // 
            this.OpenFullReportButton.Location = new System.Drawing.Point(34, 19);
            this.OpenFullReportButton.Name = "OpenFullReportButton";
            this.OpenFullReportButton.Size = new System.Drawing.Size(192, 23);
            this.OpenFullReportButton.TabIndex = 0;
            this.OpenFullReportButton.Text = "Выбрать файл полного отчета";
            this.OpenFullReportButton.UseVisualStyleBackColor = true;
            this.OpenFullReportButton.Click += new System.EventHandler(this.OpenFullReportButton_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.FullReportPathTextBox);
            this.groupBox1.Controls.Add(this.OpenFullReportButton);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(260, 96);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Шаг 1";
            // 
            // FullReportPathTextBox
            // 
            this.FullReportPathTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.FullReportPathTextBox.Location = new System.Drawing.Point(6, 48);
            this.FullReportPathTextBox.Name = "FullReportPathTextBox";
            this.FullReportPathTextBox.ReadOnly = true;
            this.FullReportPathTextBox.Size = new System.Drawing.Size(248, 20);
            this.FullReportPathTextBox.TabIndex = 1;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.InOutReportPathTextBox);
            this.groupBox2.Controls.Add(this.OpenInOutReportButton);
            this.groupBox2.Location = new System.Drawing.Point(278, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(260, 96);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Шаг 2";
            // 
            // InOutReportPathTextBox
            // 
            this.InOutReportPathTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.InOutReportPathTextBox.Location = new System.Drawing.Point(6, 48);
            this.InOutReportPathTextBox.Name = "InOutReportPathTextBox";
            this.InOutReportPathTextBox.ReadOnly = true;
            this.InOutReportPathTextBox.Size = new System.Drawing.Size(248, 20);
            this.InOutReportPathTextBox.TabIndex = 1;
            // 
            // OpenInOutReportButton
            // 
            this.OpenInOutReportButton.Location = new System.Drawing.Point(24, 19);
            this.OpenInOutReportButton.Name = "OpenInOutReportButton";
            this.OpenInOutReportButton.Size = new System.Drawing.Size(213, 23);
            this.OpenInOutReportButton.TabIndex = 0;
            this.OpenInOutReportButton.Text = "Выбрать файл отчета входа-выхода";
            this.OpenInOutReportButton.UseVisualStyleBackColor = true;
            this.OpenInOutReportButton.Click += new System.EventHandler(this.OpenInOutReportButton_Click);
            // 
            // FireButton
            // 
            this.FireButton.Location = new System.Drawing.Point(235, 188);
            this.FireButton.Name = "FireButton";
            this.FireButton.Size = new System.Drawing.Size(75, 23);
            this.FireButton.TabIndex = 3;
            this.FireButton.Text = "Огонь";
            this.FireButton.UseVisualStyleBackColor = true;
            this.FireButton.Click += new System.EventHandler(this.FireButton_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.CheckSchedule);
            this.groupBox3.Location = new System.Drawing.Point(142, 114);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(260, 68);
            this.groupBox3.TabIndex = 2;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Шаг 3";
            // 
            // CheckSchedule
            // 
            this.CheckSchedule.Location = new System.Drawing.Point(34, 19);
            this.CheckSchedule.Name = "CheckSchedule";
            this.CheckSchedule.Size = new System.Drawing.Size(192, 23);
            this.CheckSchedule.TabIndex = 0;
            this.CheckSchedule.Text = "Проверить график";
            this.CheckSchedule.UseVisualStyleBackColor = true;
            this.CheckSchedule.Click += new System.EventHandler(this.CheckSchedule_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(549, 223);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.FireButton);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "Main";
            this.Text = "AngryBoss";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button OpenFullReportButton;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox FullReportPathTextBox;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox InOutReportPathTextBox;
        private System.Windows.Forms.Button OpenInOutReportButton;
        private System.Windows.Forms.Button FireButton;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button CheckSchedule;
    }
}