﻿using System;
using System.IO;
using System.Windows.Forms;

namespace ExelSample
{
    public partial class Main : Form
    {
        private Agregator agregator;
        public Main()
        {
            InitializeComponent();
            Properties.Settings.Default.Path = Path.GetDirectoryName(Application.ExecutablePath) + @"\\data\\template" +
                                   ".rtf";
            //Properties.Settings.Default.Extention;
            Properties.Settings.Default.Save();
            agregator = new Agregator();
            ProgressBarForm progressBarForm = new ProgressBarForm();
            agregator.onSend += progressBarForm.ChangeProgress; //подписка
        }

        private void OpenFullReportButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog chooseFile = new OpenFileDialog
            {
                Filter = "Excel files|*.xls"
            };

            if (chooseFile.ShowDialog() == DialogResult.OK)
            {
                FullReportPathTextBox.Text = chooseFile.FileName;
                agregator.fullReportPath = chooseFile.FileName;
                label1.Text = chooseFile.SafeFileName;
            }
        }

        private void OpenInOutReportButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog chooseFile = new OpenFileDialog
            {
                Filter = "Excel files|*.xls*"
            };

            if (chooseFile.ShowDialog() == DialogResult.OK)
            {
                InOutReportPathTextBox.Text = chooseFile.FileName;
                label2.Text = chooseFile.SafeFileName;
                agregator.inOutReportPath = chooseFile.FileName;
            }
        }

        private void FireButton_Click(object sender, EventArgs e)
        {
            if (InOutReportPathTextBox.Text != String.Empty && FullReportPathTextBox.Text != string.Empty)
            {
                //Thread thread = new Thread(RunThread);
                //thread.Start();

                bool result = agregator.ReadAndParse(); //производится чтение из файлов и парсится в список agregator.employees
                if (result)
                {
                    agregator.FindChiefForLatecomers();  //для опоздавших находятся начальники
                    //thread.Join();
                    //выводим список для проверки информации
                    if (agregator.CheckNoID())
                    {
                        MessageBox.Show("Внимание!Обнаружены сотрудники без табельного номера!");
                        EmployeesWithoutID employeesWithoutId = new EmployeesWithoutID(agregator);
                        employeesWithoutId.ShowDialog(this);
                    }
                    LatecomersTable latecomersTable = new LatecomersTable(agregator.employees);
                    latecomersTable.ShowDialog(this);

                    //и отправляем
                    if (latecomersTable.NeedSent)
                    {
                        if (MessageBox.Show("Вы действительно хотите осуществить рассылку?", "Подтверждение",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            agregator.SendMessages();
                            MessageBox.Show("Отправка завершена");
                        }
                    }


                }
            }
            else MessageBox.Show("Чего то не хватает","Ошибка!");
        }

        private void CheckSchedule_Click(object sender, EventArgs e)
        {
            Schedule schedule = new Schedule(agregator);
            schedule.ShowDialog(this);
        }

        private void SettingsButton_Click(object sender, EventArgs e)
        {
            Settings settings = new Settings(agregator);
            settings.ShowDialog(this);
        }

        private void ShowLatecomers_Click(object sender, EventArgs e)
        {
            LatecomersTable latecomersTable = new LatecomersTable(agregator.employees);
            latecomersTable.ShowDialog(this);
        }

      private void ChooseChiefEmailButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog chooseFile = new OpenFileDialog();
            chooseFile.Filter = "Excel files|*.xls*";

            if (chooseFile.ShowDialog() == DialogResult.OK)
            {
                ChooseChiefEmailCheckTextbox.Text = chooseFile.FileName;
                label3.Text = chooseFile.SafeFileName;
                agregator.chiefEmailsPath = chooseFile.FileName;
            }
        }

        private void Main_Load(object sender, EventArgs e)
        {
            
        }

       

        //private void RunThread()
        //{
        //    agregator.ReadAndParse();
        //    agregator.FindChiefForLatecomers();
        //}
    }
}
