using System;
using System.Windows.Forms;

namespace ExelSample
{
    public partial class Settings : Form
    {
        public Settings()
        {
            InitializeComponent();
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
    }
}
