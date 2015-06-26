using System;
using System.IO;
using System.Windows.Forms;

namespace ExelSample
{
    public partial class Settings : Form
    {
        private Agregator _agrLink;
        public Settings(Agregator _agrLink)
        {
            InitializeComponent();
            this._agrLink = _agrLink;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Email = textBox1.Text;
            Properties.Settings.Default.Password = textBox2.Text;
            Properties.Settings.Default.Port = textBox3.Text;
            Properties.Settings.Default.SMTP = textBox4.Text;
            Properties.Settings.Default.Save();
            Close();
        }

        private void Settings_Load(object sender, EventArgs e)
        {

        }

        private void LoadTemplateButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog chooseFile = new OpenFileDialog();
            chooseFile.Filter = "Word files (*rtf)|*.rtf|Word files(*.doc*)|*.doc*";

            if (chooseFile.ShowDialog() == DialogResult.OK)
            {
                _agrLink.wordTemplatePath = chooseFile.FileName;

                //сохранение шаблона
                FileInfo fileInf = new FileInfo(chooseFile.FileName);
                if (fileInf.Exists)
                {
                    Properties.Settings.Default.Extention = fileInf.Extension;
                    Properties.Settings.Default.Path = Path.GetDirectoryName(Application.ExecutablePath) + @"\\data\\template" +
                                  Properties.Settings.Default.Extention;
                    fileInf.CopyTo(Properties.Settings.Default.Path, true);
                    Properties.Settings.Default.Save();
                }
            }
        }
    }
}
