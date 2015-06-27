using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel;
using Microsoft.Office.Interop.Excel;

namespace ExelSample
{
    public partial class EmployeesWithoutID : Form
    {
        private List<ExcelLine> employessWithoutId;
        public EmployeesWithoutID(Agregator agregator)
        {
            InitializeComponent();
            ReadInOutFileNoExcel(agregator.inOutReportPath);
            ShowEmployeesWithoutID();
        }

        private void ReadInOutFileNoExcel(object path)
        {
            try
            { 
                FileStream stream = File.Open((string)path, FileMode.Open, FileAccess.Read);

                string pattern = @"\w+\.xlsx";
                Regex rg = new Regex(pattern);

                IExcelDataReader excelReader;
                excelReader = rg.IsMatch((string)path) ? ExcelReaderFactory.CreateOpenXmlReader(stream) : ExcelReaderFactory.CreateBinaryReader(stream);


                DataSet result = excelReader.AsDataSet();
                List<ExcelLine> inOutReportFile = new List<ExcelLine>();

                for (int j = 0; j < 5; j++) excelReader.Read();

                while (excelReader.Read())
                {
                    ExcelLine inOutReportLine = new ExcelLine();
                    inOutReportLine.Cell = new string[14];

                    for (int i = 0; i < 14; i++)
                    {
                        inOutReportLine.Cell[i] = excelReader.GetString(i);
                    }

                    if (string.IsNullOrEmpty(inOutReportLine.Cell[3]))
                        if (string.IsNullOrEmpty(inOutReportLine.Cell[5])) break; //достигли конца списка

                    if (inOutReportLine.Cell[2] == "")
                        inOutReportFile.Add(inOutReportLine);
                }
                excelReader.Close();
                employessWithoutId = inOutReportFile;
            }
            catch (Exception error)
            {
                MessageBox.Show(string.Format("Подробности:\n {0}\n\n{1}", error.InnerException, error.Message), "Ошибка!");
            }
        }

        private void ShowEmployeesWithoutID()
        {
            try
            {
                EmployeesWithoutIdDataGridView.ColumnCount = 2;
                EmployeesWithoutIdDataGridView.Columns[0].Width = 40;
                EmployeesWithoutIdDataGridView.Columns[1].Width = 260;

                //шапка таблицы
                EmployeesWithoutIdDataGridView.Columns[0].Name = "№ п/п";
                EmployeesWithoutIdDataGridView.Columns[1].Name = "Имя сотрудника";

                //Теперь выведем в datagrid
                for (int i = 0; i < employessWithoutId.Count; i++)
                    EmployeesWithoutIdDataGridView.Rows.Add();

                int rowNumber = 0;
                foreach (var worker in employessWithoutId)
                {
                    EmployeesWithoutIdDataGridView.Rows[rowNumber].Cells[1].Value = worker.Cell[3];
                    EmployeesWithoutIdDataGridView.Rows[rowNumber].Cells[0].Value = rowNumber + 1;
                    rowNumber++;
                }
                EmployeesWithoutIdDataGridView.AllowUserToAddRows = false;
            }
            catch (Exception error)
            {
                MessageBox.Show(string.Format("Подробности:\n {0}\n\n{1}", error.InnerException, error.Message), "Ошибка!");
            }
        }

        private void ConfirmButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int i = 0;
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application
            {
                DefaultSaveFormat = XlFileFormat.xlExcel8,
                DisplayAlerts = false
            };
            //Книга.
            Workbook ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
            //Таблица.
            Worksheet ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];

            ObjWorkSheet.Cells[1, 1].ColumnWidth = 6;
            ObjWorkSheet.Cells[1, 2].ColumnWidth = 20;
            ObjWorkSheet.Cells[1, 3].ColumnWidth = 37;

            ObjWorkSheet.Cells[1, 1] = "№ п/п";
            ObjWorkSheet.Cells[1, 2] = "№ п/п в файле отчета";
            ObjWorkSheet.Cells[1, 3] = "ФИО";

            foreach (var worker in employessWithoutId)
            {
                i++;
                ObjWorkSheet.Cells[i + 1, 1] = i;
                ObjWorkSheet.Cells[i + 1, 2] = worker.Cell[1];
                ObjWorkSheet.Cells[i + 1, 3] = worker.Cell[3];
            }

            SaveFileDialog dialog = new SaveFileDialog()
            {
                Filter = "Excel files|*.xls",
                FileName = "Список сотрудников без тн",
                AddExtension = true,
                DefaultExt = ".xls"
            };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileInfo fileInf = new FileInfo(dialog.FileName);
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
