using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Excel;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ExelSample
{
    struct ExcelLine
    {
        public string[] Cell;
    }

    public class Parser
    {
        public delegate void ProgressBarInc(int max);

        public event ProgressBarInc onParse;

        private string[,] _inOutReport;
        private string[,] _fullReport;
        private List<ExcelLine> _chiefEmaiList;
        public bool Read(string inOutReportPath, string fullReportPath, string chiefEmailsPath)
        {
            #region чтение в 2 потока (не используется)

            //Thread readInOutFileThread = new Thread(ReadInOutFile);
            //readInOutFileThread.Start(inOutReportPath);

            //Thread readFullReportFileThread = new Thread(ReadFullReport);
            //readFullReportFileThread.Start(fullReportPath);

            //readInOutFileThread.Join();
            //readFullReportFileThread.Join();

            //Marshal.CleanupUnusedObjectsInCurrentContext(); // выгрузить неуправляемые ресурсы
            //GC.Collect();

            #endregion

            #region чтение в 1 поток

            int result = 0;
            result += ReadInOutFileNoExcel(inOutReportPath);
            result += ReadFullReportNoExcel(fullReportPath);
            result += ReadCheifEmailFileNoExcel(chiefEmailsPath);
            return result >= 3;
            
            #endregion
        }

        #region Старый вариант чтения

        private void ReadInOutFile(object path)
        {
            Application objWorkExcel = new Application();
            Workbook objWorkBook = objWorkExcel.Workbooks.Open((string)path,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing); //открыть файл

            Worksheet objWorkSheet = (Worksheet)objWorkBook.Sheets[1]; //получить 1 лист
            var lastCell = objWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell); //1 ячейку
            _inOutReport = new string[lastCell.Column, lastCell.Row]; // массив значений с листа равен по размеру листу

            for (int i = 0; i < _inOutReport.GetLength(0); i++) //по всем колонкам
                for (int j = 0; j < _inOutReport.GetLength(1); j++) // по всем строкам
                    _inOutReport[i, j] = objWorkSheet.Cells[j + 1, i + 1].Text.ToString(); //считываем текст в строку

            objWorkBook.Close(false, Type.Missing, Type.Missing);
            objWorkExcel.Quit(); // выйти из экселя
            Marshal.FinalReleaseComObject(objWorkExcel);
            Marshal.CleanupUnusedObjectsInCurrentContext();
        }

        private void ReadFullReport(object path)
        {
            Application objWorkExcel = new Application();
            Workbook objWorkBook = objWorkExcel.Workbooks.Open((string)path,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing); //открыть файл
            Worksheet objWorkSheet = (Worksheet)objWorkBook.Sheets[1]; //получить 1 лист
            var lastCell = objWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell); //1 ячейку
            _fullReport = new string[15, lastCell.Row]; // массив значений с листа равен по размеру листу

            for (int i = 0; i < 15; i++) //по всем колонкам
                for (int j = 0; j < _fullReport.GetLength(1); j++) // по всем строкам
                    _fullReport[i, j] = objWorkSheet.Cells[j + 1, i + 1].Text.ToString(); //считываем текст в строку

            objWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            objWorkExcel.Quit(); // выйти из экселя
            Marshal.FinalReleaseComObject(objWorkExcel);
            Marshal.CleanupUnusedObjectsInCurrentContext();
        }

        #endregion

        private int ReadInOutFileNoExcel(object path)
        {
            FileStream stream = null;
            try
            {
                stream = File.Open((string) path, FileMode.Open, FileAccess.Read);
            }
            catch (IOException)
            {
                MessageBox.Show("Возникла ошибка. Закройте файл " + (string)path +" и повторите попытку.");
                return 0;
            }

            string pattern = @"\w+\.xlsx";
            Regex rg = new Regex(pattern);

            IExcelDataReader excelReader;
            if (rg.IsMatch((string)path))
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream); 
            else
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);

            DataSet result = excelReader.AsDataSet();       
            List<ExcelLine> inOutReportFile = new List<ExcelLine>();

            for (int j = 0; j < 4; j++)
            {
                ExcelLine inOutReportLine = new ExcelLine();
                inOutReportLine.Cell = new string[14];

                for (int i = 0; i < 14; i++)
                {
                    inOutReportLine.Cell[i] = excelReader.GetString(i);
                }

                inOutReportFile.Add(inOutReportLine);
                excelReader.Read();
            }

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

                inOutReportFile.Add(inOutReportLine);
            }
            excelReader.Close();

            _inOutReport = new string[14, inOutReportFile.Count];
            for (int i = 0; i < 14; i++)
            {
                for (int j = 0; j < inOutReportFile.Count; j++)
                {
                    _inOutReport[i, j] = inOutReportFile.ElementAt(j).Cell[i] ?? "";
                }
            }
            return 1;
        }

        private int ReadFullReportNoExcel(object path)
        {
            FileStream stream = null;
            try
            {
                stream = File.Open((string)path, FileMode.Open, FileAccess.Read);
            }
            catch (IOException)
            {
                MessageBox.Show("Возникла ошибка. Закройте файл " + (string)path + " и повторите попытку.");
                return 0;
            }

            string pattern = @"\w+\.xlsx";
            Regex rg = new Regex(pattern);

            IExcelDataReader excelReader;
            if (rg.IsMatch((string)path))
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            else
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);

            DataSet result = excelReader.AsDataSet();
            List<ExcelLine> fullReportFile = new List<ExcelLine>();

            while (excelReader.Read())
            {
                ExcelLine fullReportLine = new ExcelLine();
                fullReportLine.Cell = new string[15];

                for (int i = 0; i < 15; i++)
                {
                    fullReportLine.Cell[i] = excelReader.GetString(i);
                }

                if (string.IsNullOrEmpty(fullReportLine.Cell[0])) break; //достигли конца списка

                fullReportFile.Add(fullReportLine);
            }
            excelReader.Close();

            _fullReport = new string[15, fullReportFile.Count];
            for (int i = 0; i < 15; i++)
            {
                for (int j = 0; j < fullReportFile.Count; j++)
                {
                    _fullReport[i, j] = fullReportFile.ElementAt(j).Cell[i] ?? "";
                }
            }
            return 1;
        }

        private int ReadCheifEmailFileNoExcel(object path)
        {
            FileStream stream = null;
            try
            {
                stream = File.Open((string) path, FileMode.Open, FileAccess.Read);
            }
            catch (IOException)
            {
                MessageBox.Show("Возникла ошибка. Закройте файл " + (string)path +" и повторите попытку.");
                return 0;
            }
            string pattern = @"\w+\.xlsx";
            Regex rg = new Regex(pattern);

            IExcelDataReader excelReader;
            if (rg.IsMatch((string)path))
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            else
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);

            DataSet result = excelReader.AsDataSet();
            _chiefEmaiList = new List<ExcelLine>();

            while (excelReader.Read())
            {
                ExcelLine fullReportLine = new ExcelLine();
                fullReportLine.Cell = new string[4];

                for (int i = 0; i < 4; i++)
                {
                    fullReportLine.Cell[i] = excelReader.GetString(i);
                }

                if (string.IsNullOrEmpty(fullReportLine.Cell[0])) break; //достигли конца списка

                _chiefEmaiList.Add(fullReportLine);
            }
            excelReader.Close();
            return 1;
        }

        public List<Employee> Parse(Agregator agrLink)
        {
            List<Employee> employees = new List<Employee>();
            int numEmployees = GetNumberOfEmployees();
            agrLink.peroid = _inOutReport[0, 2];

            for (int i = 0; i < numEmployees; i++)
            {
                onParse(numEmployees); //для progressbar

                string[] fullName;
                int position = i + 2; // индекс служащего в _fullReport
                int id = Convert.ToInt32(_fullReport[5, position]); //получаем отдельно id.
                string[] subdivision; // массив подразделений
                string[] fullData = GetFullData(position, out subdivision); // полная инфрмация о сотруднике
                InOutTime[] times = GetInOutTimeData(id, agrLink); // массив, содержащий информацию о времени прихода-ухода
                string email;

                try
                {
                    email = _chiefEmaiList.Last(s => s.Cell[1].Equals(id.ToString())).Cell[3];
                }
                catch
                {
                    email = string.Empty;
                }
                if (fullData != null)
                    employees.Add(new Employee(id, fullData[0], fullData[1], fullData[2], times, fullData[3],
                        fullData[4], subdivision, email));
                else
                {
                    fullName = ParseFullName(position);
                    employees.Add(new Employee(id, fullName[0], fullName[1], fullName[2], times));
                }
            }
            //MessageBox.Show(employees.Count.ToString());
            return employees;
        }

        private int GetNumberOfEmployees()
        {
            return Int16.Parse(_fullReport[0, _fullReport.GetLength(1) - 1]);
            //upperNumber = 0;
            //numEmployees = 0;
            //startNumber = 0;
            //string pattern = @"[0-9]+";
            //Regex reg = new Regex(pattern);
            //for (int i = _inOutReport.GetLength(1) - 1; i >= 0; i--)
            //{
            //    if (reg.IsMatch(_inOutReport[1, i]))
            //    {
            //        upperNumber = i;
            //        numEmployees = Convert.ToInt32(_inOutReport[1, i]);
            //        break;
            //    }
            //}

            //for (int i = 0; i < _inOutReport.GetLength(1); i++)
            //{
            //    if (reg.IsMatch(_inOutReport[1, i]))
            //    {
            //        startNumber = i;
            //        break;
            //    }
            //}
        }

        private string[] ParseFullName(int position)
        {
            string[] fullName = new string[3];
            string pattern = @"(\w+)";
            Regex reg = new Regex(pattern);
            Match match = reg.Match(_inOutReport[3, position]);
            for (int i = 0; i < 3; i++)
            {
                fullName[i] = match.Value;
                match = match.NextMatch();
            }
            return fullName;
        }

        private string[] GetFullData(int position, out string[] division)
        {
            int j = 0;
            string[] fullData = new string[5];
            division = new string[6];

            fullData[0] = _fullReport[2, position];
            fullData[1] = _fullReport[3, position];
            fullData[2] = _fullReport[4, position];
            fullData[3] = _fullReport[13, position];
            fullData[4] = _fullReport[14, position];

            for (int i = 0; i < 6; i++)
                division[i] = _fullReport[i + 7, position];

            return fullData;
        }

        private InOutTime[] GetInOutTimeData(int id, Agregator agrLink)
        {
            InOutTime[] data = new InOutTime[7];
            int index;
            try
            {
                index = Enumerable.Range(0, _inOutReport.GetLength(1))
                    .Where(i => _inOutReport[2, i].Equals(id.ToString()))
                    .ToList().Last();
            }
            catch (SystemException ex)
            {
                return null;
            }

            string pattern = @"([0-9]+:[0-9]+:[0-9]+)|(нет)";
            Regex reg = new Regex(pattern);
            string[] temp;

            for (int i = 0; i < 7; i++)
            {
                temp = new string[2];
                int j = 0;
                Match match = reg.Match(_inOutReport[i + 5, index]);
                while (match.Success)
                {
                    temp[j] = match.Value;
                    match = match.NextMatch();
                    j++;
                }
                data[i] = new InOutTime(_inOutReport[i + 5, 4], temp[0], temp[1], agrLink);

            }
            return data;
        }

        public bool CheckNoId()
        {
            for (int i = 3; i < _inOutReport.GetLength(0); i++)
            {
                if (_inOutReport[i, 2] == String.Empty) return true;
            }
            return false;
        }
    }
}