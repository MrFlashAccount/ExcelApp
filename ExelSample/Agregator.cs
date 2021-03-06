﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using MailMessage = System.Net.Mail.MailMessage;

namespace ExelSample
{
    public class Agregator
    {
        // ReSharper disable InconsistentNaming
        public delegate void ProgressBarInc(int max);
        public delegate void ProgressBarClose();

        /// <summary>
        /// событие отправки сообщений, для progress bar
        /// </summary>
        public event ProgressBarInc onSend;

        /// <summary>
        /// событие происходящее при ошибке, по нему должен закрываться progressbar
        /// </summary>
        public event ProgressBarClose onError;

        /// <summary>
        /// время начала р.д. по умолчанию
        /// </summary>
        public TimeSpan StartWorkingTime = TimeSpan.Parse("8:00:00");

        /// <summary>
        /// время окончания р.д. по умолчанию
        /// </summary>
        public TimeSpan EndWorkingTime = TimeSpan.Parse("17:00:00");

        /// <summary>
        /// Расписание начала р.д.
        /// </summary>
        public Dictionary<int, TimeSpan> StartWorkingWeek = new Dictionary<int, TimeSpan>();

        /// <summary>
        /// Расписание окончания р.д.
        /// </summary>
        public Dictionary<int, TimeSpan> EndWorkingWeek = new Dictionary<int, TimeSpan>();

        /// <summary>
        /// путь к файлу входа-выхода
        /// </summary>
        public string inOutReportPath;

        /// <summary>
        /// путь к файлу полного отчета
        /// </summary>
        public string fullReportPath;

        /// <summary>
        /// путь к файлу со списком e-mail начальников
        /// </summary>
        public string chiefEmailsPath;

        /// <summary>
        /// путь к шаблону Word
        /// </summary>
        public string wordTemplatePath;
        
        /// <summary>
        /// путь к особому списку адресатов 
        /// </summary>
        public string specialEmailListPath;

        /// <summary>
        /// список сотрудников
        /// </summary>
        public List<Employee> employees;


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

        /// <summary>
        /// Чтение и разбор входных файлов
        /// </summary>
        /// <returns>true - успешно, false - нет</returns>
        public bool ReadAndParse()
        {
            parser = new Parser();
            ProgressBarForm progressBarForm = new ProgressBarForm();
            progressBarForm.Show();
            parser.onParse += progressBarForm.ChangeProgress; //подписка

            bool result = parser.Read(inOutReportPath, fullReportPath, chiefEmailsPath, specialEmailListPath);
            if (!result)
            {
                progressBarForm.Close();
                return false;
            }
            employees = parser.Parse(this);
            return true;
        }


        public bool CheckNoID()
        {
            return parser.CheckNoId();
        }

        /// <summary>
        /// Поиск начальников для опоздавших
        /// </summary>
        public void FindChiefForLatecomers()
        {
            List<Employee> chiefs = GetChiefList();
            foreach (var latecomer in employees.Where(latecomer => latecomer.IsLatest))
            {
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
        public List<Employee> GetLatecomers()
        {
            return employees.Where(s => s.IsLatest && s.NeedToSent).ToList();
        }

        /// <summary>
        /// сохранение писем локально в виде вордовских документов
        /// </summary>
        /// <returns></returns>
        public void SendLocal()
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Properties.Settings.Default.Number=1;
                    int Number = 1;
                    List<Employee> chiefs = GetChiefList();
                    List<Employee> latecomerList = GetLatecomers();

                    foreach (var latecomer in latecomerList)
                        latecomer.FindChief(chiefs);

                    foreach (Employee chief in chiefs)
                    {
                        try
                        {
                            if (onSend != null) onSend(chiefs.Count);

                            latecomersOfchief = latecomerList.Where(s => s.Chief.Id == chief.Id).ToList();
                            if (latecomersOfchief.Count != 0)
                            {
                                string path = dialog.SelectedPath + @"\\Сообщение о приходе-уходе " + (Number++) + ".rtf";
                                WordDocument report = CreateReportFromTemplate(latecomersOfchief);
                                report.Save(path);
                                report.Close();
                                Marshal.CleanupUnusedObjectsInCurrentContext();
                                GC.Collect();
                                Properties.Settings.Default.Number++;
                            }
                        }
                        catch (Exception error)
                        {
                            if (onError != null) onError();
                            MessageBox.Show("Подробности:\n " + error.InnerException + "\n\n" + error.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    MessageBox.Show("Сообщения успешно сохранены");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка!" + ex.InnerException);
                }

            }
            
        }

        /// <summary>
        /// Отправка сообщений начальникам
        /// </summary>
        public void SendMessages()
        {
            Properties.Settings.Default.Number = 1;
            List<Employee> chiefs = GetChiefList();
            List<Employee> latecomerList = GetLatecomers();
            string path = wordTemplatePath.Substring(0, wordTemplatePath.IndexOf(".", StringComparison.Ordinal)) + "temp" + Properties.Settings.Default.Extention;

            foreach (var latecomer in latecomerList)
            {
                latecomer.FindChief(chiefs);
            }
            ReplaceFromSpecialList(latecomerList);


            SmtpClient smtp = new SmtpClient(Properties.Settings.Default.SMTP, int.Parse(Properties.Settings.Default.Port))
            {
                Credentials =
                    new NetworkCredential(Properties.Settings.Default.Email, Properties.Settings.Default.Password), // входим в учетные данные
                EnableSsl = Properties.Settings.Default.SSL
                //Timeout = 999999999
            };

            //var test = new List<Employee>();
            //foreach (var s in latecomerList)
            //{
            //    if (s.Position.Contains("Заместитель") && s.Subdivision[3] == "") test.Add(s);
            //}

            int messageNumber = 0;
            foreach (Employee chief in chiefs)
            {
                try
                {
                    if (messageNumber >= 9)
                    {
                        messageNumber = 0;
                        smtp.Dispose();
                        smtp = null;
                        GC.Collect();

                        smtp = new SmtpClient(Properties.Settings.Default.SMTP, int.Parse(Properties.Settings.Default.Port))    //новый сеанс чтобы не сработала блокировка спама
                        {
                            Credentials =
                                new NetworkCredential(Properties.Settings.Default.Email, Properties.Settings.Default.Password),
                            EnableSsl = Properties.Settings.Default.SSL
                        };
                    }

                    if (onSend != null) onSend(chiefs.Count);

                    latecomersOfchief = latecomerList.Where(s => s.Chief.Id == chief.Id).ToList();
                    if (latecomersOfchief.Count != 0)
                    {
                        WordDocument report = CreateReportFromTemplate(latecomersOfchief);
                        report.Save(path);
                        report.Close();
                        Marshal.CleanupUnusedObjectsInCurrentContext();

                        if (report.Closed)
                        {
                            SendMessage(path, smtp, chief);
                            messageNumber++;
                        }
                        GC.Collect();
                        Properties.Settings.Default.Number++;
                    }  
                }
                catch (Exception error)
                {
                    if (onError != null) onError();
                    MessageBox.Show("Подробности:\n " + error.InnerException + "\n\n" + error.Message, "Ошибка!",MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            Properties.Settings.Default.Save();
            MessageBox.Show("Отправка завершена");
        }

        /// <summary>
        /// Отправка одного сообщения
        /// </summary>
        /// <param name="path">Путь к вложению</param>
        /// <param name="smtp">Smtp</param>
        /// <param name="chief">Начальник которому отправляем</param>
        private void SendMessage(string path,SmtpClient smtp,Employee chief)
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
        }

        /// <summary>
        /// Создание вложения для письма на основе шаблона
        /// </summary>
        /// <param name="latecomersOfchief"></param>
        /// <returns></returns>
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
                foreach (var s in latecomersOfchief[i].TimeList.Where(s => s.IsLatest))
                {
                    if (s.IncomeTime != null)
                        if (s.OutcomeTime != null)
                            wordDoc.Selection.Text += s.Date.ToString("dd.MM.yyyy") + ":\n" +
                                                     "пришел: " + s.IncomeTime.Value + "\n" +
                                                     "ушел: " + s.OutcomeTime.Value + "\n";
                        else
                            wordDoc.Selection.Text += s.Date.ToString("dd.MM.yyyy") + ":\n" +
                                                     "пришел: " + s.IncomeTime.Value + "\n" +
                                                     "ушел: " + "нет данных\n";
                    else if (s.OutcomeTime != null)
                        wordDoc.Selection.Text += s.Date.ToString("dd.MM.yyyy") + ":\n" +
                                                 "пришел: " + "нет данных" + "\n" + "ушел: " +
                                                 s.OutcomeTime.Value + "\n";
                    else
                        wordDoc.Selection.Text += s.Date.ToString("dd.MM.yyyy") + ":\n" + "пришел: " + "нет данных" + "\n" + "ушел: " +
                                                 "нет данных\n";
                }
            }

            #endregion

            #region заполнение оставшихся полей

            try
            {
                wordDoc.SetSelectionToBookmark("date"); //переход к закладке date
                wordDoc.Selection.Aligment = TextAligment.Center;
                wordDoc.Selection.Text = DateTime.Today.Day.ToString("D2") + "." + DateTime.Today.Month.ToString("D2") + "." + DateTime.Today.Year + "г.";  // заполнение текущей даты
            } 
            catch 
            {
                //ignored
            }
            try
            {
                wordDoc.SetSelectionToBookmark("num"); //переход к закладке num
                wordDoc.Selection.Text = Properties.Settings.Default.Number.ToString(); //заполнение текущего номера отчета
            } 
            catch
            {
                //ignored
            }
            try
            {
                wordDoc.SetSelectionToBookmark("period"); //переход к закладке period
                wordDoc.Selection.Text = peroid.Remove(0, peroid.IndexOf(" ", StringComparison.Ordinal) + 1);
            }
            catch
            {
                //ignored
            }
            try
            {
                wordDoc.SetSelectionToBookmark("name"); //переход к закладке name
                wordDoc.Selection.Aligment = TextAligment.Center;
                wordDoc.Selection.Text = "(" + latecomersOfchief[0].Chief.Surname + " " +
                                         latecomersOfchief[0].Chief.Name + " " + latecomersOfchief[0].Chief.Patronymic + ") ";
            }
            catch
            {
                //ignored
            }

            #endregion

            return wordDoc;
        }

        /// <summary>
        /// заменяем начальника если сотрудник в специальном списке
        /// </summary>
        public void ReplaceFromSpecialList()
        {
            try
            {
                foreach (var worker in employees)
                {
                    int index = parser.specialEmailList.FindIndex(p => p.Cell[0] == worker.Id.ToString());

                    if (index > -1)
                    {
                        if (GetChiefList()
                                .FindIndex(s => s.Id == Convert.ToInt32(parser.specialEmailList[index].Cell[3])) > -1)
                        {
                            worker.Chief = null;
                            worker.Chief =
                                GetChiefList()
                                    .FirstOrDefault(s => s.Id == Convert.ToInt32(parser.specialEmailList[index].Cell[3]));
                        }
                        else
                        {
                            MessageBox.Show("Ошибка особого списка: не найден руководитель с табельным номером " +
                                            Convert.ToInt32(parser.specialEmailList[index].Cell[3]));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при применении особого списка. " + ex.InnerException);
            }
        }

        public void ReplaceFromSpecialList(List<Employee> employees2)
        {
            try
            {
                foreach (var worker in employees2)
                {
                    int index = parser.specialEmailList.FindIndex(p => p.Cell[0] == worker.Id.ToString());

                    if (index > -1)
                    {
                        if (GetChiefList()
                                .FindIndex(s => s.Id == Convert.ToInt32(parser.specialEmailList[index].Cell[3])) > -1)
                        {
                            worker.Chief = null;
                            worker.Chief =
                                GetChiefList()
                                    .FirstOrDefault(s => s.Id == Convert.ToInt32(parser.specialEmailList[index].Cell[3]));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при применении особого списка. " + ex.InnerException);
            }
        }
    }
}
