using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Runtime.Hosting;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExelSample
{
    public class Agregator
    {
        public delegate void ProgressBarInc(int max);
        public event ProgressBarInc onSend;

        public TimeSpan StartWorkingTime = TimeSpan.Parse("8:00:00"); // время начала р.д. по умолчанию
        public TimeSpan EndWorkingTime = TimeSpan.Parse("17:00:00"); // время окончания р.д. по умолчанию

        public Dictionary<int, TimeSpan> StartWorkingWeek = new Dictionary<int, TimeSpan>(); //Расписание начала р.д.
        public Dictionary<int, TimeSpan> EndWorkingWeek = new Dictionary<int, TimeSpan>(); // Расписание окончания р.д.

        public string inOutReportPath; // путь к файлу входа-выхода
        public string fullReportPath; // путь к файлу полного отчета
        public string chiefEmailsPath; // путь к файлу со списком e-mail начальников
        public string wordTemplatePath; // путь к шаблону Word.
        public List<Employee> employees; // список сотрудников
        public List<Employee> latecomersOfchief;
        private Parser parser;
        public string peroid;


        public Agregator()
        {
            inOutReportPath = String.Empty;
            fullReportPath = String.Empty;
            wordTemplatePath = Properties.Settings.Default.Path;
            employees = new List<Employee>();
            FillDictionaries();
        }
        public bool ReadAndParse()
        {
            parser = new Parser();
            ProgressBarForm progressBarForm = new ProgressBarForm();
            progressBarForm.Show();
            parser.onParse += progressBarForm.ChangeProgress; //подписка

            bool result = parser.Read(inOutReportPath, fullReportPath, chiefEmailsPath);
            if (!result) return false;
            employees = parser.Parse(this);
            return true;
            //FindChiefForLatecomers();
            //SendMessages();
        }

        public bool CheckNoID()
        {
            return parser.CheckNoId();
        }

        public void FindChiefForLatecomers()
        {
            List<Employee> chiefs = GetChiefList();
            foreach (var latecomer in employees)
            {
                if (latecomer.IsLatest)
                    latecomer.FindChief(chiefs);
            }
        }

        /// <summary>
        /// Заполняет расписание р.д.
        /// </summary>
        private void FillDictionaries()
        {
            for (int i = 0; i < 7; i++)
            {
                if (i >= 1 && i < 5) // с понедельника по четверг
                {
                    StartWorkingWeek.Add(i, StartWorkingTime);
                    EndWorkingWeek.Add(i, EndWorkingTime);
                }
                else if (i == 5) // если пятница
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
            return  employees.Where(s => s.Category == "Руководитель" || s.Category == "Ведущий менеджер").Where(s => s.Subdivision[3] == "").ToList(); // Если руководитель и не ниже 3 уровня.
        }

        /// <summary>
        /// Возвращает список опздавших в течение недели.
        /// </summary>
        /// <returns>список опоздавших</returns>
        private List<Employee> GetLatecomers()
        {
            return employees.Where(s => s.IsLatest && s.NeedToSent).ToList();
        }

        public void SendMessages()
        {
            Properties.Settings.Default.Number++;
            List<Employee> chiefs = GetChiefList();
            List<Employee> latecomerList = GetLatecomers();

            foreach (var latecomer in latecomerList)
            {
                latecomer.FindChief(chiefs);
            }

            #region тестирование

            //var smtp = new SmtpClient("smtp.gmail.com", 587)
            //{
            //    Credentials = new NetworkCredential("sergikgarin@gmail.com", "hs,fr007"),
            //    EnableSsl = true
            //};
            //smpt.Send("sergikgarin@gmail.com", "sergikgarin@gmail.com", "test", "testbody");

            #endregion

            SmtpClient smtp = new SmtpClient(Properties.Settings.Default.SMTP, int.Parse(Properties.Settings.Default.Port))
            {
                Credentials =
                    new NetworkCredential(Properties.Settings.Default.Email, Properties.Settings.Default.Password), // входим в учетные данные
                EnableSsl = true,
                //Timeout = 999999999
            };

            foreach (Employee chief in chiefs)
            {
                try
                {
                    onSend(chiefs.Count);

                    latecomersOfchief = latecomerList.Where(s => s.Chief.Id == chief.Id).ToList();
                    if (latecomersOfchief.Count != 0)
                    {
                        string path = wordTemplatePath.Substring(0, wordTemplatePath.IndexOf(".")) + "temp" +
                                      Properties.Settings.Default.Extention;
                        WordDocument report = CreateReportFromTemplate(latecomersOfchief);
                        report.Save(path);
                        report.Close();
                        Marshal.CleanupUnusedObjectsInCurrentContext();
                        if (report.Closed)
                        {
                            Attachment attachment = new Attachment(path);
                            MailMessage message = new MailMessage
                            {
                                From = new MailAddress(Properties.Settings.Default.Email),
                                Subject = "Опоздавшиe"
                            };
                            if (chief.Email != "")
                            {
                                message.To.Add(new MailAddress(chief.Email));
                                message.Attachments.Add(attachment);
                                smtp.Send(message);
                            }
                            attachment.Dispose();
                            //attachment = null;
                            GC.Collect();
                        }
                    }
                }
                catch (Exception error)
                {
                    MessageBox.Show("Ошибка!! Подробности: " + error.Message);
                    return;
                }

            }
            Properties.Settings.Default.Save();
        }

        private WordDocument CreateReportFromTemplate(List<Employee> latecomersOfchief)
        {
            WordDocument wordDoc = null;
            try
            {
                wordDoc = new WordDocument(wordTemplatePath);
            }
            catch (Exception error)
            {
                if (wordDoc != null)
                {
                    wordDoc.Close();
                }
                throw error;
            }

            #region Шапка таблицы

            wordDoc.SetSelectionToBookmark("table");
            wordDoc.InsertTable(latecomersOfchief.Count + 1, 4);

            wordDoc.SetSelectionToCell(1, 1);
            wordDoc.Selection.Text = "Табельный номер";

            wordDoc.SetSelectionToCell(1, 2);
            wordDoc.Selection.Text = "ФИО";

            wordDoc.SetSelectionToCell(1, 3);
            wordDoc.Selection.Text = "Должность";

            wordDoc.SetSelectionToCell(1, 4);
            wordDoc.Selection.Text = "Информация по опозданиям";

            #endregion

            #region заполнение таблицы

            for (int i = 0; i < latecomersOfchief.Count; i++)
            {
                int position = i + 2;

                wordDoc.SetSelectionToCell(position, 1);
                wordDoc.Selection.Text = latecomersOfchief[i].Id.ToString();

                wordDoc.SetSelectionToCell(position, 2);
                wordDoc.Selection.Text = latecomersOfchief[i].Surname + " " + latecomersOfchief[i].Name[0] + "." +
                                         latecomersOfchief[i].Patronymic[0] + ".";

                wordDoc.SetSelectionToCell(position, 3);
                wordDoc.Selection.Text = latecomersOfchief[i].Position;

                wordDoc.SetSelectionToCell(position, 4);
                //wordDoc.Selection.Text = "";
                foreach (var s in latecomersOfchief[i].TimeList.Where(s => s.IsLatest))
                {
                    if (s.IncomeTime != null)
                        if (s.OutcomeTime != null)
                            wordDoc.Selection.Text += s.Date.ToString("dd-MM-yyyy") + ":\n" +
                                                     "пришел: " + s.IncomeTime.Value + "\n" +
                                                     "ушел: " + s.OutcomeTime.Value + "\n";
                        else
                            wordDoc.Selection.Text += s.Date.ToString("dd-MM-yyyy") + ":\n" +
                                                     "пришел: " + s.IncomeTime.Value + "\n" +
                                                     "ушел: " + "нет данных\n";
                    else if (s.OutcomeTime != null)
                        wordDoc.Selection.Text += s.Date.ToString("dd-MM-yyyy") + ":\n" +
                                                 "пришел: " + "нет данных" + "\n" + "ушел: " +
                                                 s.OutcomeTime.Value + "\n";
                    else
                        wordDoc.Selection.Text += s.Date.ToString("dd-MM-yyyy") + ":\n" + "пришел: " + "нет данных" + "\n" + "ушел: " +
                                                 "нет данных\n";
                }
            }

            #endregion

            #region заполнение оставшихся полей

            wordDoc.SetSelectionToBookmark("date"); //переход к закладке date
            wordDoc.Selection.Aligment = TextAligment.Center;
            wordDoc.Selection.Text = DateTime.Today.Day + "." + DateTime.Today.Month + "." + DateTime.Today.Year + "г.";  // заполнение текущей даты

            wordDoc.SetSelectionToBookmark("num"); //переход к закладке num
            wordDoc.Selection.Text = Properties.Settings.Default.Number.ToString(); //заполнение текущего номера отчета

            wordDoc.SetSelectionToBookmark("period"); //переход к закладке period
            wordDoc.Selection.Text = peroid.Remove(0, peroid.IndexOf(" ") + 1);

            wordDoc.SetSelectionToBookmark("name"); //переход к закладке name
            wordDoc.Selection.Aligment = TextAligment.Center;
            wordDoc.Selection.Text = latecomersOfchief[0].Chief.Surname + " " + latecomersOfchief[0].Chief.Name + " " + latecomersOfchief[0].Chief.Patronymic;

            #endregion

            return wordDoc;
        }
    }
}
