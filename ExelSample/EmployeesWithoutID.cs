using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel;

namespace ExelSample
{
    public partial class EmployeesWithoutID : Form
    {
        private List<ExcelLine> employessWithoutId;
        public EmployeesWithoutID(Agregator agregator)
        {
            InitializeComponent();
            ReadInOutFileNoExcel(agregator.inOutReportPath);
            ShowLEmployeesWithoutID();
        }

        private void ReadInOutFileNoExcel(object path)
        {
            try
            { 
                FileStream stream = File.Open((string)path, FileMode.Open, FileAccess.Read);

                string pattern = @"\w+\.xlsx";
                Regex rg = new Regex(pattern);

                IExcelDataReader excelReader;
                if (rg.IsMatch((string)path))
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                else
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);


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
                MessageBox.Show("Ошибка!! Подробности: " + error.Message);
                return;
            }
        }

        private void ShowLEmployeesWithoutID()
        {
            try
            {

                EmployeesWithoutIdDataGridView.ColumnCount = 1;
                EmployeesWithoutIdDataGridView.Columns[0].Width = 300;

                //шапка таблицы
                EmployeesWithoutIdDataGridView.Columns[0].Name = "Имя сотрудника";

                //Теперь выведем в datagrid
                for (int i = 0; i < employessWithoutId.Count; i++)
                    EmployeesWithoutIdDataGridView.Rows.Add();

                int rowNumber = 0;
                foreach (var worker in employessWithoutId)
                {
                    EmployeesWithoutIdDataGridView.Rows[rowNumber].Cells[0].Value = worker.Cell[3];
                    rowNumber++;
                }
                EmployeesWithoutIdDataGridView.AllowUserToAddRows = false;
            }
            catch (Exception error)
            {
                MessageBox.Show("Ошибка!! Подробности: " + error.Message);
                return;
            }
        }

        private void ConfirmButton_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
