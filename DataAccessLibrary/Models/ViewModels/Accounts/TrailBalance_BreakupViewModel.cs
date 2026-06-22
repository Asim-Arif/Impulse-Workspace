using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccessLibrary.Models.ViewModels.Accounts
{
    public class TrailBalance_BreakupModel
    {
        public string AccNo { get; set; }
        public string AccTitle { get; set; }
        public string SubAccOf { get; set; }
        public bool Parent { get; set; }
        public bool Active { get; set; }

        public double OpeningDr { get; set; }
        public double OpeningCr { get; set; }

        public double Debit { get; set; }
        public double Credit { get; set; }

        public double ClosingDr { get; set; }
        public double ClosingCr { get; set; }

        public double AdjDr { get; set; }
        public double AdjCr { get; set; }

        //public double FinalClosingDr => ClosingDr + AdjDr;
        //public double FinalClosingCr => ClosingCr + AdjCr;

        public double FinalClosingDr { get; set; }
        public double FinalClosingCr { get; set; }

        public double OpeningBalance { get; set; }
        public double OpeningBalanceDr { get; set; }
        public double OpeningBalanceCr { get; set; }
        public double LastBalanceBalanceDr { get; set; }
        public double LastBalanceBalanceCr { get; set; }

    }
}