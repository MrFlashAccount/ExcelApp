using System.Collections.Generic;
using System.Linq;

// ReSharper disable InconsistentNaming

namespace ExelSample
{
    public class Employee
    {
        #region Обявление параметров класса

        public int Id { get; private set; } //Идентификатор
        public string Surname { get; private set; } //Фамилия
        public string Name { get; private set; } //Имя
        public string Patronymic { get; private set; } //Отчество
        public string Position { get; private set; } //Должность
        public string Category { get; private set; } //Категория
        public string[] Subdivision { get; private set; } //Массив подразделений
        public List<InOutTime> TimeList { get; private set; } // Список данных о времени прихода-ухода.
        public string Email { get; private set; } // адрес электронной почты
        public Employee Chief { get; private set; } //ссылка на начальника.
        public bool IsLatest { get; private set; } // Клеймо опоздавшего

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
            Chief = chiefList.Last(s => s.Subdivision[2].Equals(Subdivision[2]));
        }
    }
}