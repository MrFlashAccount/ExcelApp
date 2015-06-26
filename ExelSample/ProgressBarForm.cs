﻿using System.Windows.Forms;

namespace ExelSample
{
    public partial class ProgressBarForm : Form
    {
        public ProgressBarForm()
        {
            InitializeComponent();
            progressBar1.Value = 0;
        }

        public void ChangeProgress(int max)
        {
            if (progressBar1.Value == 1) Show();
            progressBar1.Maximum = max;
            progressBar1.Increment(1);
            if (progressBar1.Value >= max) Close();
        }

        private void progressBar1_Click(object sender, System.EventArgs e)
        {

        }
    }
}
