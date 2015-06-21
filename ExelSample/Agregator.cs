using System;
using System.Collections.Generic;

namespace ExelSample
{
    public class Agregator
    {
        public TimeSpan StartWorkingTime = TimeSpan.Parse("8:00:00");
        public TimeSpan EndWorkingTime = TimeSpan.Parse("17:00:00");

        public Dictionary<int, TimeSpan> StartWorkingWeek = new Dictionary<int, TimeSpan>();
        public Dictionary<int, TimeSpan> EndWorkingWeek = new Dictionary<int, TimeSpan>();

        public string inOutReportPath;
        public string fullReportPath;
        public List<Employee> employees;


        public Agregator()
        {
            inOutReportPath = String.Empty;
            fullReportPath = String.Empty;
            employees = new List<Employee>();
            FillDictionaries();
        }
        public void ReadAndParse()
        {
            Parser parser = new Parser();
            parser.Read(inOutReportPath, fullReportPath);
            employees = parser.Parse(this);
        }

        private void FillDictionaries()
        {
            for (int i = 0; i < 7; i++)
            {
                if (i >= 0 && i < 4)
                {
                    StartWorkingWeek.Add(i, StartWorkingTime);
                    EndWorkingWeek.Add(i, EndWorkingTime);
                }
                else if (i == 4)
                {
                    StartWorkingWeek.Add(i, StartWorkingTime);
                    EndWorkingWeek.Add(i, EndWorkingTime.Subtract(new TimeSpan(1, 0, 0)));
                }
                else
                {
                    StartWorkingWeek.Add(i, new TimeSpan(0, 0, 0));
                    EndWorkingWeek.Add(i, new TimeSpan(0, 0, 0));
                }
            }
        }
    }
}
