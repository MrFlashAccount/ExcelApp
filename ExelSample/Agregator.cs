using System;
using System.Collections.Generic;

namespace ExelSample
{
    public class Agregator
    {
        public string inOutReportPath;
        public string fullReportPath;
        public List<Employee> employees;

        public Agregator()
        {
            inOutReportPath = String.Empty;
            fullReportPath = String.Empty;
            employees = new List<Employee>();
        }
        public void ReadAndParse()
        {
            Parser parser = new Parser();
            parser.Read(inOutReportPath, fullReportPath);
            employees = parser.Parse();
        }
    }
}
