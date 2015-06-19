using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExelSample
{
    public struct InOutTime
    {
        public string IncomeTime { get; set; }
        public string OutcomeTime { get; set; }

        public InOutTime(string incomeTime, string outcomeTime) : this()
        {
            this.IncomeTime = incomeTime;
            this.OutcomeTime = outcomeTime;
        }
    }
}