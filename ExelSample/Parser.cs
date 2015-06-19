using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace ExelSample
{
    public class Parser
    {
        private string[,] _inOutReport;
        private string[,] _fullReport;
        public void Read(string inOutReportPath,string fullReportPath)
        {
            #region чтение в 2 потока

            Thread readInOutFileThread = new Thread(ReadInOutFile);
            readInOutFileThread.Start(inOutReportPath);

            Thread readFullReportFileThread = new Thread(ReadFullReport);
            readFullReportFileThread.Start(fullReportPath);

            readInOutFileThread.Join();
            readFullReportFileThread.Join();

            #endregion

            #region чтение в 1 поток

            //ReadInOutFile(inOutReportPath);
            //ReadFullReport(fullReportPath);

            #endregion
            Marshal.CleanupUnusedObjectsInCurrentContext();
            GC.Collect();
        }

        private void ReadInOutFile(object path)
        {
            Excel.Application objWorkExcel = new Excel.Application();
            Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open((string)path,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing); //открыть файл

            Excel.Worksheet objWorkSheet = (Excel.Worksheet)objWorkBook.Sheets[1]; //получить 1 лист
            var lastCell = objWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell); //1 ячейку
            _inOutReport = new string[lastCell.Column, lastCell.Row]; // массив значений с листа равен по размеру листу

            for (int i = 0; i < _inOutReport.GetLength(0); i++) //по всем колонкам
                for (int j = 0; j < _inOutReport.GetLength(1); j++) // по всем строкам
                    _inOutReport[i, j] = objWorkSheet.Cells[j + 1, i + 1].Text.ToString(); //считываем текст в строку

            objWorkBook.Close(false, Type.Missing, Type.Missing);
            objWorkExcel.Quit(); // выйти из экселя
            Marshal.FinalReleaseComObject(objWorkExcel);
            Marshal.CleanupUnusedObjectsInCurrentContext();
            GC.Collect(); // убрать за собой
        }

        private void ReadFullReport(object path)
        {
            Excel.Application objWorkExcel = new Excel.Application();
            Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open((string)path,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing); //открыть файл
            Excel.Worksheet objWorkSheet = (Excel.Worksheet)objWorkBook.Sheets[1]; //получить 1 лист
            var lastCell = objWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell); //1 ячейку
            _fullReport = new string[lastCell.Column, lastCell.Row]; // массив значений с листа равен по размеру листу

            for (int i = 0; i < _fullReport.GetLength(0); i++) //по всем колонкам
                for (int j = 0; j < _fullReport.GetLength(1); j++) // по всем строкам
                    _fullReport[i, j] = objWorkSheet.Cells[j + 1, i + 1].Text.ToString(); //считываем текст в строку

            objWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            objWorkExcel.Quit(); // выйти из экселя
            Marshal.FinalReleaseComObject(objWorkExcel);
            Marshal.CleanupUnusedObjectsInCurrentContext();
            GC.Collect(); // убрать за собой
        }

        public List<Employee> Parse()
        {
            List<Employee> employees = new List<Employee>();
            int numEmployees;
            int upperNumber;
            int startNumber;
            int count = 0;

            GetNumberOfEmployees(out numEmployees, out upperNumber, out startNumber);
            for (int i = 0; i < numEmployees; i++)
            {
                string[] fullName;
                int id;
                if (_inOutReport[2, i + startNumber] != String.Empty)
                {
                    id = Convert.ToInt32(_inOutReport[2, i + startNumber]);
                    string[] division;
                    string[] fullData = GetFullData(id, out division);
                    if (fullData != null)
                    {
                        employees.Add(new Employee(id, fullData[0], fullData[1], fullData[2], fullData[3], fullData[4], division));
                        count++;
                    }
                    else
                    {
                        fullName = ParseFullName(i + startNumber);
                        employees.Add(new Employee(id, fullName[0], fullName[1], fullName[2]));
                    }
                }
                else
                {
                    id = 0;
                    fullName = ParseFullName(i + startNumber);
                    employees.Add(new Employee(id, fullName[0], fullName[1], fullName[2]));
                }
            }
            Console.WriteLine(count);
            return employees;
        }

        private void GetNumberOfEmployees(out int numEmployees, out int upperNumber, out int startNumber)
        {
            upperNumber = 0;
            numEmployees = 0;
            startNumber = 0;
            string pattern = @"[0-9]+";
            Regex reg = new Regex(pattern);
            for (int i = _inOutReport.GetLength(1) - 1; i >= 0; i--)
            {
                if (reg.IsMatch(_inOutReport[1, i]))
                {
                    upperNumber = i;
                    numEmployees = Convert.ToInt32(_inOutReport[1, i]);
                    break;
                }
            }

            for (int i = 0; i < _inOutReport.GetLength(1); i++)
            {
                if (reg.IsMatch(_inOutReport[1, i]))
                {
                    startNumber = i;
                    break;
                }
            }
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

        private string[] GetFullData(int id, out string[] division)
        {
            int index;
            int j = 0;
            string[] fullData = new string[5];
            division = new string[6];

            var newV = Enumerable.Range(0, _fullReport.GetLength(1)).Where(i => _fullReport[5, i].Equals(id.ToString())).ToArray();
            if (newV.Length > 0)
                index = newV[0];
            else return null;

            fullData[0] = _fullReport[2, index];
            fullData[1] = _fullReport[3, index];
            fullData[2] = _fullReport[4, index];
            fullData[3] = _fullReport[13, index];
            fullData[4] = _fullReport[14, index];

            while(j < 6)
            {
                division[j] = _fullReport[j + 7, index];
                j++;
            }

            return fullData;
        }
    }
}