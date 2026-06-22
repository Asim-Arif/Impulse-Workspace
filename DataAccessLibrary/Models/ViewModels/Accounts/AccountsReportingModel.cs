using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccessLibrary.Models.ViewModels.Accounts
{
    public class AccountsReportingModel
    {
        public string AccNo { get; set; }
        public string AccTitle { get; set; }
        public DateTime DTFrom { get; set; }
        public DateTime DTTo { get; set; }
        public decimal OpeningBalance { get; set; }
        public decimal Balance { get; set; }
        public DateTime VDate { get; set; }
        public string VchrNo { get; set; }
        public string Description { get; set; }
        public Decimal Debit { get; set; }
        public Decimal Credit { get; set; }
        public string ChqNo { get; set; }
        public DateTime ChqDate { get; set; }
        public DateTime BillDate { get; set; }
        public string BillNo { get; set; }
        public string GeneratedBy { get; set; }
        public string MachineName { get; set; }
        public string UserName { get; set; }
        public DateTime EntryDate { get; set; }
        public string EditedBy { get; set; }
        public string EditMachineName { get; set; }
        public DateTime EditedDate { get; set; }
        public string VoucherType { get; set; }
        public string ChqType { get; set; }
        public string DeleteReason { get; set; }

    }
}
