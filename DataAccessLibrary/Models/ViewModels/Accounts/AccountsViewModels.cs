using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace DataAccessLibrary.Models.ViewModels.Accounts
{
    public class ChartOfAccountsModel
    {
        public string HeadType { get; set; }
        public string HeadTypeNo { get; set; }
        public string SubAccOf { get; set; }
        public string SubAccTitle { get; set; }
        public string AccNo { get; set; }
        public string AccTitle { get; set; }
        public decimal OpeningBalance { get; set; }
        public decimal Balance { get; set; }
        public string ReportingGroup { get; set; }
        public string Budget { get; set; }
        public string Budget_G { get; set; }
        public int TaxPercentage { get; set; }
        public DateTime OpenDate { get; set; }
        public int BalType { get; set; }
        public Boolean AccType { get; set; }
        public Boolean Active { get; set; }
        public Boolean bMainhead { get; set; }


        public Boolean ShowInExpense { get; set; }
        public Boolean ShowInBs { get; set; }
        public string Notes { get; set; }
        public Boolean Marked { get; set; }
        public string GroupName { get; set; }
        public string Note { get; set; }
        public Boolean ManualMarked { get; set; }
        public Boolean ParentMarked { get; set; }
        public Boolean MarkedTB { get; set; }
        public Boolean Marked_CF { get; set; }
        public string Description { get; set; }
        public long CFHFS_RefID { get; set; }
        public Boolean Marked_CF_Nve { get; set; }
        public int NewProject { get; set; }
        public string BudgetHeadDesc { get; set; }
        public string GroupDescForBudgetHead { get; set; }
        public long BH_GFER_RefID { get; set; }
        public long AccGroup_RefID { get; set; }
    }

    public class CashFlowHeadsModel
    {
        public long EntryID { get; set; }
        public string Description { get; set; }
    }

    public class NotesToAccountsModel
    {
        public long EntryID { get; set; }
        public string GroupName { get; set; }
    }
}
