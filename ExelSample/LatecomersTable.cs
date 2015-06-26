using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace ExelSample
{
    public partial class LatecomersTable : Form
    {
        private List<Employee> employeesList;
        public bool NeedSent = true;
        public LatecomersTable(List<Employee> employees)
        {
            InitializeComponent();
            ShowLatecomers(employees);
            employeesList = employees;
        }

        private void ShowLatecomers(List<Employee> employeesList)
        {
            DataGridViewCheckBoxColumn checkColumn = new DataGridViewCheckBoxColumn
            {
                Width = 50,
                ReadOnly = false,
                FillWeight = 10
            };

            LatecomersDataGridView.Columns.Add(checkColumn);

            LatecomersDataGridView.ColumnCount = 11;
            LatecomersDataGridView.Columns[2].Width = 200;
            for(int i=3; i<9; i++)
                LatecomersDataGridView.Columns[i].Width = 120;
            LatecomersDataGridView.Columns[10].Width = 200;

            //шапка таблицы
            LatecomersDataGridView.Columns[1].Name = "Т/н";
            LatecomersDataGridView.Columns[2].Name = "ФИО";
            LatecomersDataGridView.Columns[3].Name = "Понедельник";
            LatecomersDataGridView.Columns[4].Name = "Вторник";
            LatecomersDataGridView.Columns[5].Name = "Среда";
            LatecomersDataGridView.Columns[6].Name = "Четверг";
            LatecomersDataGridView.Columns[7].Name = "Пятница";
            LatecomersDataGridView.Columns[8].Name = "Суббота";
            LatecomersDataGridView.Columns[9].Name = "Воскресенье";
            LatecomersDataGridView.Columns[10].Name = "Начальник";

            //надо сделать выборку опоздавших
            List<Employee> latecomersList = employeesList.Where(s => s.IsLatest).ToList();

            //Теперь выведем в datagrid
            for (int i = 0; i < latecomersList.Count; i++)
            {
                LatecomersDataGridView.Rows.Add();
            }
            int count = 0;
            int rowNumber = 0;
            foreach (var worker in latecomersList)
            {
                count++;
                LatecomersDataGridView.Rows[rowNumber].Cells[0].Value = true;
                LatecomersDataGridView.Rows[rowNumber].Cells[2].Value = worker.Surname + " " + worker.Name + " " + worker.Patronymic;
                for (int i = 0; i < worker.TimeList.Count; i++)
                {
                    if (worker.TimeList.ElementAt(i).IncomeTime.ToString() == "00:00:00"
                        || worker.TimeList.ElementAt(i).IncomeTime == null)
                        LatecomersDataGridView.Rows[rowNumber].Cells[3 + i].Value = "отсутствовал";
                    else
                        LatecomersDataGridView.Rows[rowNumber].Cells[3 + i].Value = 
                            worker.TimeList.ElementAt(i).IncomeTime + " - " + worker.TimeList.ElementAt(i).OutcomeTime;
                }
                LatecomersDataGridView.Rows[rowNumber].Cells[10].Value = worker.Chief.Surname + " " + worker.Chief.Name + " " + worker.Chief.Patronymic;
                LatecomersDataGridView.Rows[rowNumber].Cells[1].Value = worker.Id;
                rowNumber++;
            }
            LatecomersDataGridView.AllowUserToAddRows = false;
            //MessageBox.Show(count.ToString());
        }

        private void MarkSelected(List<Employee> employeesList)
        {
            //помечаем в списке то что отметили в datagrid
            for (int i = 0; i < LatecomersDataGridView.Rows.Count; i++)
            {
                if (Convert.ToBoolean(LatecomersDataGridView.Rows[i].Cells[0].Value)) //если помечено флажком
                {
                    employeesList.Find(p => int.Parse(LatecomersDataGridView.Rows[i].Cells[1].Value.ToString()) == p.Id)
                        .NeedToSent = true;
                }
            }
        }

        private void ConfirmButton_Click(object sender, EventArgs e)
        {
            MarkSelected(employeesList);
            Close();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            NeedSent = false;
            Close();
        }
    }
}
