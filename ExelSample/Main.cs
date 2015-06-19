using System;
using System.Windows.Forms;

namespace ExelSample
{
    public partial class Main : Form
    {
        private Agregator agregator;
        public Main()
        {
            InitializeComponent();
            agregator = new Agregator();
        }

        private void OpenFullReportButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog chooseFile = new OpenFileDialog();

            if (chooseFile.ShowDialog() == DialogResult.OK)
            {
                FullReportPathTextBox.Text = chooseFile.FileName;
            }
            agregator.fullReportPath = chooseFile.FileName;
        }

        private void OpenInOutReportButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog chooseFile = new OpenFileDialog();

            if (chooseFile.ShowDialog() == DialogResult.OK)
            {
                InOutReportPathTextBox.Text = chooseFile.FileName;
            }
            agregator.inOutReportPath = chooseFile.FileName;
        }

        private void FireButton_Click(object sender, EventArgs e)
        {
            if (InOutReportPathTextBox.Text != String.Empty && FullReportPathTextBox.Text != string.Empty)
            {
                //Thread thread = new Thread(RunThread);
                //thread.Start();
                FireButton.Enabled = false;
                OpenFullReportButton.Enabled = false;
                OpenInOutReportButton.Enabled = false;
                agregator.ReadAndParse();
                FireButton.Enabled = true;
                OpenFullReportButton.Enabled = true;
                OpenInOutReportButton.Enabled = true;
            }
            else MessageBox.Show("Чего то не хватает","Ошибка!");
        }

        //private void RunThread()
        //{
        //    agregator.ReadAndParse();
        //}
    }
}
