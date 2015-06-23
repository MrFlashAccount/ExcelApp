using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;

namespace ExelSample
{
    public class Agregator
    {
        public TimeSpan StartWorkingTime = TimeSpan.Parse("8:00:00"); // время начала р.д. по умолчанию
        public TimeSpan EndWorkingTime = TimeSpan.Parse("17:00:00"); // время окончания р.д. по умолчанию

        public Dictionary<int, TimeSpan> StartWorkingWeek = new Dictionary<int, TimeSpan>(); //Расписание начала р.д.
        public Dictionary<int, TimeSpan> EndWorkingWeek = new Dictionary<int, TimeSpan>(); // Расписание окончания р.д.

        public string inOutReportPath; // путь к файлу входа-выхода
        public string fullReportPath; // путь к файлу полного отчета
        public List<Employee> employees; // список сотрудников


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
            SendMessages();
        }
        /// <summary>
        /// Заполняет расписание р.д.
        /// </summary>
        private void FillDictionaries()
        {
            for (int i = 0; i < 7; i++)
            {
                if (i >= 0 && i < 4) // с понедельника по четверг
                {
                    StartWorkingWeek.Add(i, StartWorkingTime);
                    EndWorkingWeek.Add(i, EndWorkingTime);
                }
                else if (i == 4) // если пятница
                {
                    StartWorkingWeek.Add(i, StartWorkingTime);
                    EndWorkingWeek.Add(i, EndWorkingTime.Subtract(new TimeSpan(1, 0, 0)));
                }
                else // иначе суббота и воскресенье
                {
                    StartWorkingWeek.Add(i, new TimeSpan(0, 0, 0));
                    EndWorkingWeek.Add(i, new TimeSpan(0, 0, 0));
                }
            }
        }
        /// <summary>
        /// Возвращает список начальников не ниже 3 уровня
        /// </summary>
        /// <returns>список начальников</returns>
        private List<Employee> GetChiefList()
        {
            return  employees.Where(s => s.Category == "Руководитель").Where(s => s.Subdivision[3] == "").ToList(); // Если руководитель и не ниже 3 уровня.
        }

        /// <summary>
        /// Возвращает список опздавших в течение недели.
        /// </summary>
        /// <returns>список опоздавших</returns>
        private List<Employee> GetLatecomers()
        {
            return employees.Where(s => s.IsLatest).ToList();
        }

        public void SendMessages()
        {
            List<Employee> chiefs = GetChiefList();
            bool continueMessage = false;

            foreach (var employee in employees)
            {
                employee.FindChief(chiefs);
            }

            SmtpClient smtp = new SmtpClient("smtp.yandex.ru", Int32.Parse(Properties.Settings.Default.Port))
            {
                Credentials = new NetworkCredential(Properties.Settings.Default.Email, Properties.Settings.Default.Password),
                EnableSsl = true
            };
            // входим в учетные данные

            MailMessage message = new MailMessage
            {
                From = new MailAddress("m1026m@yandex.ru"),
                Subject = "Опоздавшиe"
            };
        }
    }
}
