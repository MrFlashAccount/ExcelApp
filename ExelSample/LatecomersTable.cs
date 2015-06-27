using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace ExelSample
{
    public partial class LatecomersTable : Form
    {
        private List<Employee> employeesList;
        public bool NeedSent = true;

        public LatecomersTable(List<Employee> employees)
        {
            InitializeComponent();
            employeesList = employees;
            ShowLatecomers();
        }

        private void ShowLatecomers()
        {
            DataGridViewCheckBoxColumn checkColumn = new DataGridViewCheckBoxColumn
            {
                Width = 50,
                ReadOnly = false,
                FillWeight = 10
            };

            LatecomersDataGridView.Columns.Add(checkColumn);

            LatecomersDataGridView.ColumnCount = 12;
            LatecomersDataGridView.Columns[2].Width = 200;
            for (int i = 3; i < 10; i++)
                LatecomersDataGridView.Columns[i].Width = 120;
            LatecomersDataGridView.Columns[11].Width = 200;

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
            LatecomersDataGridView.Columns[10].Name = "Т/н начальника";
            LatecomersDataGridView.Columns[11].Name = "Начальник";

            //надо сделать выборку опоздавших
            employeesList = employeesList.Where(s => s.IsLatest).ToList();

            //Теперь выведем в datagrid
            for (int i = 0; i < employeesList.Count; i++)
            {
                LatecomersDataGridView.Rows.Add();
            }
            int rowNumber = 0;
            foreach (var worker in employeesList)
            {
                LatecomersDataGridView.Rows[rowNumber].Cells[0].Value = true;
                LatecomersDataGridView.Rows[rowNumber].Cells[1].Value = worker.Id;
                LatecomersDataGridView.Rows[rowNumber].Cells[2].Value = worker.Surname + " " + worker.Name + " " +
                                                                        worker.Patronymic;
                for (int i = 0; i < worker.TimeList.Count; i++)
                {
                    if (worker.TimeList.ElementAt(i).IncomeTime.ToString() == "00:00:00"
                        || worker.TimeList.ElementAt(i).IncomeTime == null)
                        LatecomersDataGridView.Rows[rowNumber].Cells[3 + i].Value = "отсутствовал";
                    else
                        LatecomersDataGridView.Rows[rowNumber].Cells[3 + i].Value =
                            worker.TimeList.ElementAt(i).IncomeTime + " - " + worker.TimeList.ElementAt(i).OutcomeTime;
                }
                LatecomersDataGridView.Rows[rowNumber].Cells[10].Value = worker.Chief.Id;
                LatecomersDataGridView.Rows[rowNumber].Cells[11].Value = worker.Chief.Surname + " " + worker.Chief.Name +
                                                                         " " + worker.Chief.Patronymic;
                rowNumber++;
            }
            LatecomersDataGridView.AllowUserToAddRows = false;
        }

        private void MarkSelected()
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
            MarkSelected();
            Close();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            NeedSent = false;
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int i = 1;
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Workbook ObjWorkBook;
            Worksheet ObjWorkSheet;
            //Книга.
            ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
            //Таблица.
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet) ObjWorkBook.Sheets[1];

            ObjWorkSheet.Cells.ColumnWidth = 40;
            ObjWorkSheet.Cells[1, 1].ColumnWidth = 6;
            ObjWorkSheet.Cells[1, 2].ColumnWidth = 12;

            ObjWorkSheet.Cells[1, 1] = "№ п/п";
            ObjWorkSheet.Cells[1, 2] = "Т/н";
            ObjWorkSheet.Cells[1, 3] = "ФИО";
            ObjWorkSheet.Cells[1, 4] = "Понедельник";
            ObjWorkSheet.Cells[1, 5] = "Вторник";
            ObjWorkSheet.Cells[1, 6] = "Среда";
            ObjWorkSheet.Cells[1, 7] = "Четверг";
            ObjWorkSheet.Cells[1, 8] = "Пятница";
            ObjWorkSheet.Cells[1, 9] = "Суббота";
            ObjWorkSheet.Cells[1, 10] = "Воскресенье";
            ObjWorkSheet.Cells[1, 11] = "Т/н начальника";
            ObjWorkSheet.Cells[1, 12] = "Начальник";

            foreach (var worker in employeesList)
            {
                i++;
                ObjWorkSheet.Cells[i, 1] = (i - 1).ToString();
                ObjWorkSheet.Cells[i, 2] = worker.Id.ToString();
                ObjWorkSheet.Cells[i, 3] = worker.Surname + " " + worker.Name + " " + worker.Patronymic;
                for (int j = 0; j < worker.TimeList.Count; j++)
                {
                    if (worker.TimeList.ElementAt(j).IncomeTime.ToString() == "00:00:00"
                        || worker.TimeList.ElementAt(j).IncomeTime == null)
                        ObjWorkSheet.Cells[i, 4 + j] = "отсутствовал";
                    else
                        ObjWorkSheet.Cells[i, 4 + j] = worker.TimeList.ElementAt(j).IncomeTime + " - " + worker.TimeList.ElementAt(j).OutcomeTime;
                }
                ObjWorkSheet.Cells[i, 11] = worker.Chief.Id;
                ObjWorkSheet.Cells[i, 12] =  worker.Chief.Surname + " " + worker.Chief.Name + " " + worker.Chief.Patronymic;
            }
            SaveFileDialog dialog = new SaveFileDialog()
            {
                Filter = "Excel files|*.xls",
                FileName = "Отчет об опоздавших",
                AddExtension = true,
                DefaultExt = ".xls"
            };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileInfo fileInf = new FileInfo(dialog.FileName);
                    ObjExcel.DefaultSaveFormat = XlFileFormat.xlExcel8;
                    ObjWorkBook.SaveAs(fileInf.FullName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing);
                    ObjWorkBook.Close();
                    button1.Text = "выгружено";
                    button1.Enabled = false;
                }
                catch (Exception ex)
                {
                    if (MessageBox.Show(
                        "Возникла ошибка про сохранении файла. Вы хотите попытаться отобразить его в Excel?\n" +
                        "Подробности:\n " + ex.InnerException + "\n\n" + ex.Message, "Ошибка",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Error) == DialogResult.Yes)
                    {
                        try
                        {
                            ObjExcel.Visible = true;
                            ObjExcel.UserControl = true;
                            button1.Text = "выгружено";
                            button1.Enabled = false;
                        }
                        catch (Exception error)
                        {
                            MessageBox.Show("Подробности:\n " + error.InnerException + "\n\n" + error.Message, "Ошибка!",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                            ObjWorkBook.Close();
                            Marshal.CleanupUnusedObjectsInCurrentContext();
                            return;
                        }
                    }
                    else
                    {
                        ObjWorkBook.Close();
                        Marshal.CleanupUnusedObjectsInCurrentContext();
                    }
                }
            }
            else if (MessageBox.Show("Файл не будет сохранен. Может открыть его в Excel?", "Вопрос",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                ObjExcel.Visible = true;
                ObjExcel.UserControl = true;
                button1.Text = "выгружено";
                button1.Enabled = false;
            }
            Marshal.CleanupUnusedObjectsInCurrentContext();
        }
    }
}
