using System;
using System.Collections.Generic;
using System.Linq;

// ReSharper disable InconsistentNaming

namespace ExelSample
{
    public class Employee
    {
        #region Обявление параметров класса

        public int Id { get; private set; } //Идентификатор
        public string Surname { get; set; } //Фамилия
        public string Name { get; set; } //Имя
        public string Patronymic { get; set; } //Отчество
        public string Position { get; private set; } //Должность
        public string Category { get; private set; } //Категория
        public string[] Subdivision { get; private set; } //Массив подразделений
        public List<InOutTime> TimeList { get; private set; } // Список данных о времени прихода-ухода
        public string Email { get; set; } // адрес электронной почты
        public Employee Chief { get; set; } //ссылка на начальника.
        public bool IsLatest { get; private set; } // Клеймо опоздавшего
        public bool NeedToSent { get; set; } //нужно ли отправлять

        #endregion

        public Employee(int Id, string Surname, string Name, string Patronymic, InOutTime[] times = null, string Position = "", string Category = "", string[] Subdivision = null, string Email = "")
        {
            this.Id = Id; 
            this.Surname = Surname;
            this.Name = Name;
            this.Patronymic = Patronymic;
            this.Position = Position;
            this.Category = Category;
            this.Subdivision = new string[6];
            this.Email = Email;
            NeedToSent = false;
            if (Subdivision != null)
                for (int i = 0; i < 6; i++)  this.Subdivision[i] = Subdivision[i];
            TimeList = new List<InOutTime>();
            if (times != null)
                foreach (InOutTime i in times)
                {
                    TimeList.Add(i);
                    if (i.IsLatest) //если за неделю опоздал-ставим метку опоздавшего
                        IsLatest = true;
                }
        }

        public void FindChief(List<Employee> chiefList)
        {
            if ((Category == "Руководитель" || Category == "Ведущий менеджер") && Subdivision[3] == String.Empty)
            {
                if (Position.Contains("аместитель"))
                    try
                    {
                        Chief = chiefList.Last(s => s.Subdivision[2].Equals(Subdivision[2]) && !s.Position.Contains("аместитель"));
                    }
                    catch (InvalidOperationException)
                    {
                        Chief = chiefList.Last(s => s.Id == 14001333);
                    }
                else
                    Chief = chiefList.Last(s => s.Id == 14001333);
            }
            else
                try
                {
                    Chief = chiefList.Last(s => s.Subdivision[2].Equals(Subdivision[2]));
                }
                catch (InvalidOperationException)
                {
                    Chief = chiefList.Last(s => s.Id == 14001333);
                }
            //if ((Category == "Руководитель" || Category == "Ведущий менеджер") && Subdivision[3] == String.Empty &&
            //    !Position.Contains("Заместитель"))
            //{
            //    Chief = chiefList.Last(s => s.Id == 14001333);
            //}
            //else
            //    try
            //    {
            //        Chief = chiefList.Last(s => s.Subdivision[2].Equals(Subdivision[2]));
            //    }
            //    catch (InvalidOperationException)
            //    {
            //        Chief = chiefList.Last(s => s.Id == 14001333);
            //    }

            //if ((Category == "Руководитель" || Category == "Ведущий менеджер") && Position.Contains("Заместитель"))
            //{
            //    try
            //    {
            //        Chief = chiefList.Last(s => s.Subdivision[2].Equals(Subdivision[2]));
            //    }
            //    catch (InvalidOperationException)
            //    {
            //        Chief = chiefList.Last(s => s.Id == 14001333);
            //    }
            //}
        }
    }
}