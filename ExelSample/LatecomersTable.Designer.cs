﻿namespace ExelSample
{
    partial class LatecomersTable
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
            this.LatecomersDataGridView = new System.Windows.Forms.DataGridView();
            this.ConfirmButton = new System.Windows.Forms.Button();
            this.CancelButton = new System.Windows.Forms.Button();
            this.ExportToExcel = new System.Windows.Forms.Button();
            this.SaveLocal = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.LatecomersDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // LatecomersDataGridView
            // 
            this.LatecomersDataGridView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.LatecomersDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.LatecomersDataGridView.Location = new System.Drawing.Point(12, 12);
            this.LatecomersDataGridView.Name = "LatecomersDataGridView";
            this.LatecomersDataGridView.Size = new System.Drawing.Size(808, 414);
            this.LatecomersDataGridView.TabIndex = 0;
            // 
            // ConfirmButton
            // 
            this.ConfirmButton.Location = new System.Drawing.Point(542, 435);
            this.ConfirmButton.Name = "ConfirmButton";
            this.ConfirmButton.Size = new System.Drawing.Size(136, 23);
            this.ConfirmButton.TabIndex = 1;
            this.ConfirmButton.Text = "Начать отправку";
            this.ConfirmButton.UseVisualStyleBackColor = true;
            this.ConfirmButton.Click += new System.EventHandler(this.ConfirmButton_Click);
            // 
            // CancelButton
            // 
            this.CancelButton.Location = new System.Drawing.Point(684, 435);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(136, 23);
            this.CancelButton.TabIndex = 2;
            this.CancelButton.Text = "Отмена";
            this.CancelButton.UseVisualStyleBackColor = true;
            this.CancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // ExportToExcel
            // 
            this.ExportToExcel.Location = new System.Drawing.Point(12, 435);
            this.ExportToExcel.Name = "ExportToExcel";
            this.ExportToExcel.Size = new System.Drawing.Size(140, 23);
            this.ExportToExcel.TabIndex = 3;
            this.ExportToExcel.Text = "Выгрузить в Excel";
            this.ExportToExcel.UseVisualStyleBackColor = true;
            this.ExportToExcel.Click += new System.EventHandler(this.ExportToExcel_Click);
            // 
            // SaveLocal
            // 
            this.SaveLocal.Location = new System.Drawing.Point(158, 435);
            this.SaveLocal.Name = "SaveLocal";
            this.SaveLocal.Size = new System.Drawing.Size(136, 23);
            this.SaveLocal.TabIndex = 4;
            this.SaveLocal.Text = "Сохранить сообщения";
            this.SaveLocal.UseVisualStyleBackColor = true;
            this.SaveLocal.Click += new System.EventHandler(this.SaveLocal_Click);
            // 
            // LatecomersTable
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(832, 470);
            this.Controls.Add(this.SaveLocal);
            this.Controls.Add(this.ExportToExcel);
            this.Controls.Add(this.CancelButton);
            this.Controls.Add(this.ConfirmButton);
            this.Controls.Add(this.LatecomersDataGridView);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "LatecomersTable";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Список опоздавших";
            ((System.ComponentModel.ISupportInitialize)(this.LatecomersDataGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView LatecomersDataGridView;
        private System.Windows.Forms.Button ConfirmButton;
        private System.Windows.Forms.Button CancelButton;
        private System.Windows.Forms.Button ExportToExcel;
        private System.Windows.Forms.Button SaveLocal;
    }
}