using System;
using System.Collections.Generic;

namespace ExelSample
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Employee> employees = new List<Employee>();
            Parser parser = new Parser();
            parser.Read(@"C:\Users\Сергей\Desktop\Практика\Hours.xls", @"C:\Users\Сергей\Desktop\Практика\33_Polny_otchet_16_06_2015.xls");
            employees = parser.Parse();
            Console.ReadKey();
        }
    }
}
