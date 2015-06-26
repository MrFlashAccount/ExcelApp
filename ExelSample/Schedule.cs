using System;
using System.Collections.Generic;
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

            //запишем воскресенье из-за дурацкой американской недели
            int j = 0;
            if (_agrLink.StartWorkingWeek.Values.ElementAt(j).Hours != 0)
                dataGridView1.Rows[0].Cells[j + 6].Value =
                    _agrLink.StartWorkingWeek.Values.ElementAt(j).Hours.ToString("D2") + ":" +
                    _agrLink.StartWorkingWeek.Values.ElementAt(j).Minutes.ToString("D2");
            else dataGridView1.Rows[0].Cells[j + 6].Value = "выходной";

            if (_agrLink.StartWorkingWeek.Values.ElementAt(j).Hours != 0)
                dataGridView1.Rows[1].Cells[j + 6].Value =
                    _agrLink.EndWorkingWeek.Values.ElementAt(j).Hours.ToString("D2") + ":" +
                    _agrLink.EndWorkingWeek.Values.ElementAt(j).Minutes.ToString("D2");
            else dataGridView1.Rows[1].Cells[j + 6].Value = "выходной";

            for (int i = 1; i < 7; i++)
            {
                if (_agrLink.StartWorkingWeek.Values.ElementAt(i).Hours != 0)
                    dataGridView1.Rows[0].Cells[i - 1].Value =
                        _agrLink.StartWorkingWeek.Values.ElementAt(i).Hours.ToString("D2") + ":" +
                        _agrLink.StartWorkingWeek.Values.ElementAt(i).Minutes.ToString("D2");
                else dataGridView1.Rows[0].Cells[i - 1].Value = "выходной";

                if (_agrLink.StartWorkingWeek.Values.ElementAt(i).Hours != 0)
                    dataGridView1.Rows[1].Cells[i - 1].Value =
                        _agrLink.EndWorkingWeek.Values.ElementAt(i).Hours.ToString("D2") + ":" +
                        _agrLink.EndWorkingWeek.Values.ElementAt(i).Minutes.ToString("D2");
                else dataGridView1.Rows[1].Cells[i - 1].Value = "выходной";
            }
        }

        /// <summary>
        /// Сохранение изменений в расписании
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ConfirmButton_Click(object sender, EventArgs e)
        {
            Dictionary<int, TimeSpan> StartWorkingWeek = new Dictionary<int, TimeSpan>(); //Расписание начала р.д.
            Dictionary<int, TimeSpan> EndWorkingWeek = new Dictionary<int, TimeSpan>(); // Расписание окончания р.д.

            //для воскресенья
            int j = 0;
            StartWorkingWeek.Add(j, (dataGridView1.Rows[0].Cells[j + 6].Value.ToString() == "выходной") ?
                new TimeSpan(0, 0, 0) :
                TimeSpan.Parse(dataGridView1.Rows[0].Cells[j + 6].Value.ToString()));

            EndWorkingWeek.Add(j, (dataGridView1.Rows[1].Cells[j + 6].Value.ToString() == "выходной") ?
                new TimeSpan(0, 0, 0) :
                TimeSpan.Parse(dataGridView1.Rows[1].Cells[j + 6].Value.ToString()));

            for (int i = 1; i < 7; i++)
            {
                StartWorkingWeek.Add(i, (dataGridView1.Rows[0].Cells[i - 1].Value.ToString() == "выходной") ?
                    new TimeSpan(0, 0, 0) :
                    TimeSpan.Parse(dataGridView1.Rows[0].Cells[i - 1].Value.ToString()));

                EndWorkingWeek.Add(i, (dataGridView1.Rows[1].Cells[i - 1].Value.ToString() == "выходной") ?
                    new TimeSpan(0, 0, 0) :
                    TimeSpan.Parse(dataGridView1.Rows[1].Cells[i - 1].Value.ToString()));
            }

            _agrLink.StartWorkingWeek = StartWorkingWeek;
            _agrLink.EndWorkingWeek = EndWorkingWeek;

            Close();
        }

        /// <summary>
        /// Выход после просмотра расписания
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CancelButton_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
