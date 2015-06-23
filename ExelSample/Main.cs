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
                agregator.fullReportPath = chooseFile.FileName;
                label1.Text = chooseFile.SafeFileName;
            }
        }

        private void OpenInOutReportButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog chooseFile = new OpenFileDialog();

            if (chooseFile.ShowDialog() == DialogResult.OK)
            {
                InOutReportPathTextBox.Text = chooseFile.FileName;
                label2.Text = chooseFile.SafeFileName;
                agregator.inOutReportPath = chooseFile.FileName;
            }
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

        private void CheckSchedule_Click(object sender, EventArgs e)
        {
            Schedule schedule = new Schedule(agregator);
            schedule.ShowDialog(this);
        }

        private void SettingsButton_Click(object sender, EventArgs e)
        {
            Settings settings = new Settings();
            settings.ShowDialog(this);
        }

        //private void RunThread()
        //{
        //    agregator.ReadAndParse();
        //}
    }
}
