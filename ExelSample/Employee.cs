using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace ExelSample
{
    public class Employee
    {
        #region Обявление параметров класса

        public int Id { get; set; }
        public string Surname { get; set; }
        public string Name { get; set; }
        public string Patronymic { get; set; }
        public string Position { get; set; }
        public string Category { get; set; }
        public string[] Subdivision { get; set; }
        public List<InOutTime> TimeList { get; set; }

        #endregion

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        public Employee(int Id, string Surname, string Name, string Patronymic, string Position = "", string Category = "", string[] Subdivision = null, InOutTime[] times = null)
        {
            this.Id = Id;
            this.Surname = Surname;
            this.Name = Name;
            this.Patronymic = Patronymic;
            this.Position = Position;
            this.Category = Category;
            this.Subdivision = new string[6];
            if (Subdivision != null)
                for (int i = 0; i < 6; i++)  this.Subdivision[i] = Subdivision[i];
            if (times != null)
                foreach (InOutTime i in times)
                    TimeList.Add(i);
        }

        public void Write()
        {
            for (int i = 0; i < 6; i++)
            {
                Console.Write(Subdivision[i] + " ");
            }
        }

    }
}