using System.ComponentModel;
using System.Windows.Forms;

namespace ExelSample
{
    partial class Main
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main));
            this.FireButton = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.SettingsButton = new System.Windows.Forms.Button();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.ChooseChiefEmailCheckTextbox = new System.Windows.Forms.TextBox();
            this.ChooseChiefEmailButton = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.CheckSchedule = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.FullReportPathTextBox = new System.Windows.Forms.TextBox();
            this.OpenFullReportButton = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.InOutReportPathTextBox = new System.Windows.Forms.TextBox();
            this.OpenInOutReportButton = new System.Windows.Forms.Button();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.groupBox4.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.SuspendLayout();
            // 
            // FireButton
            // 
            this.FireButton.BackColor = System.Drawing.Color.IndianRed;
            this.FireButton.Location = new System.Drawing.Point(46, 19);
            this.FireButton.Name = "FireButton";
            this.FireButton.Size = new System.Drawing.Size(159, 23);
            this.FireButton.TabIndex = 3;
            this.FireButton.Text = "Огонь";
            this.FireButton.UseVisualStyleBackColor = false;
            this.FireButton.Click += new System.EventHandler(this.FireButton_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // SettingsButton
            // 
            this.SettingsButton.Location = new System.Drawing.Point(416, 231);
            this.SettingsButton.Name = "SettingsButton";
            this.SettingsButton.Size = new System.Drawing.Size(122, 23);
            this.SettingsButton.TabIndex = 5;
            this.SettingsButton.Text = "Настройки";
            this.SettingsButton.UseVisualStyleBackColor = true;
            this.SettingsButton.Click += new System.EventHandler(this.SettingsButton_Click);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.label3);
            this.groupBox4.Controls.Add(this.ChooseChiefEmailCheckTextbox);
            this.groupBox4.Controls.Add(this.ChooseChiefEmailButton);
            this.groupBox4.Location = new System.Drawing.Point(12, 114);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(260, 92);
            this.groupBox4.TabIndex = 11;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Шаг 3";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 71);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(0, 13);
            this.label3.TabIndex = 3;
            // 
            // ChooseChiefEmailCheckTextbox
            // 
            this.ChooseChiefEmailCheckTextbox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.ChooseChiefEmailCheckTextbox.Location = new System.Drawing.Point(6, 48);
            this.ChooseChiefEmailCheckTextbox.Name = "ChooseChiefEmailCheckTextbox";
            this.ChooseChiefEmailCheckTextbox.ReadOnly = true;
            this.ChooseChiefEmailCheckTextbox.Size = new System.Drawing.Size(248, 20);
            this.ChooseChiefEmailCheckTextbox.TabIndex = 1;
            // 
            // ChooseChiefEmailButton
            // 
            this.ChooseChiefEmailButton.Location = new System.Drawing.Point(16, 19);
            this.ChooseChiefEmailButton.Name = "ChooseChiefEmailButton";
            this.ChooseChiefEmailButton.Size = new System.Drawing.Size(228, 23);
            this.ChooseChiefEmailButton.TabIndex = 0;
            this.ChooseChiefEmailButton.Text = "Выбрать файл со списком начальников";
            this.ChooseChiefEmailButton.UseVisualStyleBackColor = true;
            this.ChooseChiefEmailButton.Click += new System.EventHandler(this.ChooseChiefEmailButton_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.CheckSchedule);
            this.groupBox3.Location = new System.Drawing.Point(278, 114);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(260, 92);
            this.groupBox3.TabIndex = 9;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Шаг 4";
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
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.FullReportPathTextBox);
            this.groupBox1.Controls.Add(this.OpenFullReportButton);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(260, 96);
            this.groupBox1.TabIndex = 8;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Шаг 1";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 71);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 13);
            this.label1.TabIndex = 2;
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
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.InOutReportPathTextBox);
            this.groupBox2.Controls.Add(this.OpenInOutReportButton);
            this.groupBox2.Location = new System.Drawing.Point(278, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(260, 96);
            this.groupBox2.TabIndex = 10;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Шаг 2";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 71);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(0, 13);
            this.label2.TabIndex = 3;
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
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.FireButton);
            this.groupBox5.Location = new System.Drawing.Point(150, 212);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(245, 55);
            this.groupBox5.TabIndex = 12;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Шаг 5";
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(550, 279);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.SettingsButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Main";
            this.Text = "AngryBoss";
            this.Load += new System.EventHandler(this.Main_Load);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private Button FireButton;
        private OpenFileDialog openFileDialog1;
        private Button SettingsButton;
        private GroupBox groupBox4;
        private Label label3;
        private TextBox ChooseChiefEmailCheckTextbox;
        private Button ChooseChiefEmailButton;
        private GroupBox groupBox3;
        private Button CheckSchedule;
        private GroupBox groupBox1;
        private Label label1;
        private TextBox FullReportPathTextBox;
        private Button OpenFullReportButton;
        private GroupBox groupBox2;
        private Label label2;
        private TextBox InOutReportPathTextBox;
        private Button OpenInOutReportButton;
        private GroupBox groupBox5;
    }
}