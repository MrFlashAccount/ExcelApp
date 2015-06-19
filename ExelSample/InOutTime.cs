namespace ExelSample
{
    public struct InOutTime
    {
        public string IncomeTime { get; set; }
        public string OutcomeTime { get; set; }
        public string Date { get; set; }

        public InOutTime(string date,string incomeTime = "", string outcomeTime = "") : this()
        {
            IncomeTime = incomeTime;
            OutcomeTime = outcomeTime;
            Date = date;
        }
    }
}