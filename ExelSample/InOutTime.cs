using System;
using System.Linq;

namespace ExelSample
{
    public struct InOutTime
    {
        public TimeSpan? IncomeTime { get; set; }
        public TimeSpan? OutcomeTime { get; set; }
        public DateTime Date { get; set; }
        public bool IsLatest { get; private set; }

        public InOutTime(string date, string incomeTime, string outcomeTime, Agregator agrLink)
            : this()
        {
            switch (incomeTime)
            {
                case null:
                case "нет":
                    IncomeTime = null;
                    break;
                case "":
                    IncomeTime = new TimeSpan(0, 0, 0);
                    break;
                default:
                    IncomeTime = TimeSpan.Parse(incomeTime);
                    break;
            }
            switch (outcomeTime)
            {
                case null:
                case "нет":
                    OutcomeTime = null;
                    break;
                case "":
                    OutcomeTime = new TimeSpan(0, 0, 0);
                    break;
                default:
                    OutcomeTime = TimeSpan.Parse(outcomeTime);
                    break;
            }
            Date = DateTime.Parse(date);
            IsLatest = CheckLatest(agrLink);
        }

        private bool CheckLatest(Agregator agrLink)
        {

            //если выходной, все в порядке
            if (agrLink.StartWorkingWeek.ElementAt((int)Date.DayOfWeek).Value.ToString() == "00:00:00")
                return false;

            if (IncomeTime == null) //если сотрудника не было
            {
                return true;
            }
            else
            {
                int compareResult =
                    IncomeTime.Value.CompareTo(agrLink.StartWorkingWeek.ElementAt((int)Date.DayOfWeek).Value);
                if (compareResult > 0) return true; //если все в порядке, проверим еще и уход
            }

            if (OutcomeTime == null) //если сотрудника не было
            {
                return true;
            }
            else
            {
                int compareResult =
                    OutcomeTime.Value.CompareTo(agrLink.StartWorkingWeek.ElementAt((int)Date.DayOfWeek).Value);
                return compareResult < 0;
            }


            //if (IncomeTime != null)
            //{
            //    int compareResult =
            //        IncomeTime.Value.CompareTo(agrLink.StartWorkingWeek.ElementAt((int) Date.DayOfWeek).Value);
            //    //Сравниваем время прихода с расписанием
            //    return compareResult > 0;
            //}
            //else
            //{
            //    //Если его не было и это был не выходной - наказать
            //    if (agrLink.StartWorkingWeek.ElementAt((int) Date.DayOfWeek).Value.ToString() != "00:00:00")
            //        return true;

            //}

            //if (OutcomeTime != null)
            //{
            //    int compareResult = OutcomeTime.Value.CompareTo(agrLink.EndWorkingWeek.ElementAt((int) Date.DayOfWeek).Value);
            //        //Сравниваем время ухода с расписанием
            //    return compareResult < 0;
            //}
            //else
            //{
            //    //Если его не было и это был не выходной - наказать
            //    if (agrLink.StartWorkingWeek.ElementAt((int)Date.DayOfWeek).Value.ToString() != "00:00:00")
            //        return true;
            //}


            //return true; //значит прошел без карточки и нужно его наказать
            ////return false;
        }
    }
}