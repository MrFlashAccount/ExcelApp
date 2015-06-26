namespace ExelSample
{
    partial class EmployeesWithoutID
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
            this.EmployeesWithoutIdDataGridView = new System.Windows.Forms.DataGridView();
            this.ConfirmButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.EmployeesWithoutIdDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // EmployeesWithoutIdDataGridView
            // 
            this.EmployeesWithoutIdDataGridView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.EmployeesWithoutIdDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.EmployeesWithoutIdDataGridView.Location = new System.Drawing.Point(8, 12);
            this.EmployeesWithoutIdDataGridView.Name = "EmployeesWithoutIdDataGridView";
            this.EmployeesWithoutIdDataGridView.Size = new System.Drawing.Size(360, 414);
            this.EmployeesWithoutIdDataGridView.TabIndex = 1;
            // 
            // ConfirmButton
            // 
            this.ConfirmButton.Location = new System.Drawing.Point(116, 436);
            this.ConfirmButton.Name = "ConfirmButton";
            this.ConfirmButton.Size = new System.Drawing.Size(123, 23);
            this.ConfirmButton.TabIndex = 2;
            this.ConfirmButton.Text = "ОК";
            this.ConfirmButton.UseVisualStyleBackColor = true;
            this.ConfirmButton.Click += new System.EventHandler(this.ConfirmButton_Click);
            // 
            // EmployeesWithoutID
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(376, 471);
            this.Controls.Add(this.ConfirmButton);
            this.Controls.Add(this.EmployeesWithoutIdDataGridView);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "EmployeesWithoutID";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "EmployeesWithoutID";
            ((System.ComponentModel.ISupportInitialize)(this.EmployeesWithoutIdDataGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView EmployeesWithoutIdDataGridView;
        private System.Windows.Forms.Button ConfirmButton;
    }
}