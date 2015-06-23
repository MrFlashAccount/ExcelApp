using System;
using System.Linq;
using System.Windows.Forms;

namespace ExelSample
{
    public partial class Schedule : Form
    {
        private readonly Agregator _agrLink;
        public Schedule(Agregator agrLink)
        {
            InitializeComponent();
            _agrLink = agrLink;
        }

        private void Schedule_Load(object sender, EventArgs e)
        {
            dataGridView1.ColumnCount = 7;
            dataGridView1.RowCount = 2;
            dataGridView1.Columns[0].Width = 100;
            dataGridView1.Columns[0].Name = "Понедельник";
            dataGridView1.Columns[1].Name = "Вторник";
            dataGridView1.Columns[2].Name = "Среда";
            dataGridView1.Columns[3].Name = "Четверг";
            dataGridView1.Columns[4].Name = "Пятница";
            dataGridView1.Columns[5].Name = "Суббота";
            dataGridView1.Columns[6].Name = "Воскресенье";

            for (int i = 0; i < 7; i++)
            {
                if (_agrLink.StartWorkingWeek.Values.ElementAt(i).Hours != 0)
                    dataGridView1.Rows[0].Cells[i].Value =
                        _agrLink.StartWorkingWeek.Values.ElementAt(i).Hours.ToString("D2") + "." +
                        _agrLink.StartWorkingWeek.Values.ElementAt(i).Minutes.ToString("D2");
                else dataGridView1.Rows[0].Cells[i].Value = "выходной";

                if (_agrLink.StartWorkingWeek.Values.ElementAt(i).Hours != 0)
                    dataGridView1.Rows[1].Cells[i].Value =
                        _agrLink.EndWorkingWeek.Values.ElementAt(i).Hours.ToString("D2") + "." +
                        _agrLink.EndWorkingWeek.Values.ElementAt(i).Minutes.ToString("D2");
                else dataGridView1.Rows[1].Cells[i].Value = "выходной";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {

            }
        }
    }
}
