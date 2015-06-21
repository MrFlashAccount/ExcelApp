using System.Collections.Generic;

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
        public bool IsLatest { get; private set; }

        #endregion

       
        public Employee(int Id, string Surname, string Name, string Patronymic, InOutTime[] times = null, string Position = "", string Category = "", string[] Subdivision = null)
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
            TimeList = new List<InOutTime>();
            if (times != null)
                foreach (InOutTime i in times)
                {
                    TimeList.Add(i);
                    if (i.IsLatest) //если за неделю опоздал-ставим метку опоздавшего
                        IsLatest = true;
                }
        }
    }
}