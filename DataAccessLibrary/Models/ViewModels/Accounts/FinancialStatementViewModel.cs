using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccessLibrary.Models.ViewModels.Accounts
{
    public class MainHeadsModel
    {
        public int EntryID { get; set; }
        public string GroupName { get; set; }
        public string SortNo { get; set; }
        public int SequenceNo { get; set; }
        public string Description { get; set; }
    }

    public class SubHeadsModel
    {
        public int EntryID { get; set; }
        public string GroupName { get; set; }
        public string MainGroupName { get; set; }
        public int Type { get; set; }
        public int SortNo { get; set; }
        public string SubDesc { get; set; }
        public string Notes { get; set; }
        public bool ShowDetail { get; set; }
        public bool bShowSubHeadDetail_NTA { get; set; }
    }

    public class SubHeadsListModel
    {
        public int EntryID { get; set; }
        public string GroupName { get; set; }
    }

    public class MainHeadsListModel
    {
        public int EntryID { get; set; }
        public string GroupName { get; set; }
    }

    public class SubHeadAccountDetailModel
    {
        public int EntryID { get; set; }
        public int AccGroup_RefID { get; set; }
        public string GroupName { get; set; } = string.Empty;
        public string AccNo { get; set; } = string.Empty;
        public string AccTitle { get; set; } = string.Empty;
        public string SubAccOf { get; set; } = string.Empty;
        public string Notes { get; set; } = string.Empty;
        public string NoteNo { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string UserName { get; set; } = string.Empty;
        public string MachineName { get; set; } = string.Empty;
        public bool Marked { get; set; }
    }
    public class GroupsAccountsModel
    {
        public int EntryID { get; set; }
        public string GroupName { get; set; }
        public string SortNo { get; set; }
        public string Type { get; set; }
        public string MainGroupName { get; set; }
        public string UserName { get; set; }
        public string MachineName { get; set; }
    }
    public class AccountsHeadManualEntryModel
    {
        public int EntryID { get; set; }
        public int AccGroup_RefID { get; set; }
        public string GroupName { get; set; }
        public string SubNotes { get; set; }
        public DateTime DT { get; set; }
        public string Description { get; set; }
        public string Notes { get; set; }
        public string Amount { get; set; }
        public bool isDetail { get; set; }
        public string UserName { get; set; }
        public string MachineName { get; set; }
    }

    public class AccountsHeadManualEntry_SC_Model
    {
        public int EntryID { get; set; }
        public DateTime DT { get; set; }
        public int AccGroup_RefID { get; set; }
        public string GroupName { get; set; }
        public string SubNotes { get; set; }
        public string Description { get; set; }
        public string Notes { get; set; }
        public double TotalAmount { get; set; }
        public double ShareValue { get; set; }
        public double ShareRate { get; set; }
        public string UserName { get; set; }
        public string MachineName { get; set; }
    }
    public class RemunerationandOtherBenefitsModel
    {
        public int EntryID { get; set; }
        public DateTime DT { get; set; }
        public int AccGroup_RefID { get; set; }
        public string GroupName { get; set; }
        public string SubNotes { get; set; }
        public string Description { get; set; }
        public string Notes { get; set; }
        public double CEAmount { get; set; }
        public double DirAmount { get; set; }
        public double ExeAmount { get; set; }
        public string UserName { get; set; }
        public string MachineName { get; set; }
    }

    public class ManualEntryForNotesModel
    {
        public int EntryID { get; set; }
        public int AccGroup_RefID { get; set; }
        public string GroupName { get; set; }
        public string SubNotes { get; set; }
        public DateTime DT { get; set; }
        public string Description { get; set; }
        public string Notes { get; set; }
        public string Amount { get; set; }
        public bool Marked { get; set; }
        public string UserName { get; set; }
        public string MachineName { get; set; }
    }

    public class SubHeadsManualEntryModel
    {
        public int EntryID { get; set; }
        public int MainAccGroup_RefID { get; set; }
        public string MainGroupName { get; set; }
        public string MainNoteNo { get; set; }
        public int AccGroup_RefID { get; set; }
        public string GroupName { get; set; }
        public string SubNotes { get; set; }
        public DateTime DTFrom { get; set; }
        public DateTime DTTo { get; set; }
        public long SrNo { get; set; }
        public string Description { get; set; }
        public string Notes { get; set; }
        public double OpeningBalance { get; set; }
        public double ClosingBalance { get; set; }
        public bool Marked { get; set; }
        public string UserName { get; set; }
        public string MachineName { get; set; }
    }

    public class Vouchers_Adjustment_Transaction_Model
    {
        public int EntryID { get; set; }
        public int RefID { get; set; }
        public DateTime DT { get; set; }
        public string VchrNo { get; set; }
        public string AccNo { get; set; }
        public string Description { get; set; }
        public double Debit { get; set; }
        public double Credit { get; set; }
        public double Balance { get; set; }
        public string Remarks { get; set; }
        public int Type { get; set; }
        public bool Marked { get; set; }
        public string SubAccOf { get; set; }
        public string AccTitle { get; set; }
        public string UserName { get; set; }
        public string MachineName { get; set; }
        public DateTime DTEntry { get; set; }
    }

    public class Vouchers_Main_Model
    {
        public long EntryID { get; set; }
        public string VchrNo { get; set; }
        public DateTime DT { get; set; }
        public string Remarks { get; set; }
        public string UserName { get; set; }
        public string MachineName { get; set; }

        public List<Vouchers_Detail_Model> LineItems { get; set; } = new List<Vouchers_Detail_Model>();

        public decimal TotalDebit => LineItems.Sum(li => li.Debit);
        public decimal TotalCredit => LineItems.Sum(li => li.Credit);
    }
    public class Vouchers_Detail_Model
    {
        public int RefID { get; set; }
        public string VchrNo { get; set; }
        public string AccNo { get; set; }
        public string Description { get; set; }
        public string AccTitle { get; set; }
        public decimal Debit { get; set; }
        public decimal Credit { get; set; }
        public decimal Balance { get; set; }
        public bool IsSystemLine { get; set; } = false; //this is for UI handling, it's not matched to database field.
    }
    public class ProfitandLossHeads_Main_Model
    {
        public long EntryID { get; set; }
        public string Description { get; set; }
        public int SrNo { get; set; }
        public bool IsTotal { get; set; }
        public bool Marked { get; set; }
        public string NoteNo { get; set; }
        public List<ProfitandLossHeads_Detail_Model> DetailItems { get; set; } = new List<ProfitandLossHeads_Detail_Model>();
    }
    public class ProfitandLossHeads_Detail_Model
    {
        public long RefID { get; set; }
        public long Group_RefID { get; set; }
        public bool IsMain { get; set; } = false; 
    }

    public class ProfitandLossHeadsDetailToLoadModel
    {
        public long EntryID { get; set; }
        public long RefID { get; set; }
        public string NoteNo { get; set; }
        public string Description { get; set; }
        public int SrNo { get; set; }
        public int Group_RefID { get; set; }
        public bool IsMain { get; set; }
        public string GroupName { get; set; }
        public string Note { get; set; }
        public bool IsTotal { get; set; }
        public bool Marked { get; set; }
    }

    public class FinancialStatementHeadsModel
    {
        public long EntryID { get; set; }
        public long RefID { get; set; }
        public string GroupName { get; set; }
        public string Note { get; set; }
        public bool MainHead { get; set; }
        public string MainNoteNo { get; set; }
        public bool ShowInBS { get; set; }
        public bool ShowInNTA { get; set; }
        public int SrNo { get; set; }

    }

    public class BalanceSheet_Main_Model
    {
        public long EntryID { get; set; }
        public string Description { get; set; }
        public int SrNo { get; set; }
        public bool IsTotal { get; set; }
        public bool Marked { get; set; }
        public string NoteNo { get; set; }
        public int AssetType { get; set; }
        public string HeadType { get; set; }
        public List<BalanceSheet_Detail_Model> DetailItems { get; set; } = new List<BalanceSheet_Detail_Model>();
    }
    public class BalanceSheet_Detail_Model
    {
        public long RefID { get; set; }
        public long Group_RefID { get; set; }
        public bool IsMain { get; set; } = false;
    }

    public class BalanceSheetHeadsDetailToLoadModel
    {
        public long EntryID { get; set; }
        public long RefID { get; set; }
        public string NoteNo { get; set; }
        public string Description { get; set; }
        public int SrNo { get; set; }
        public int Group_RefID { get; set; }
        public bool IsMain { get; set; }
        public string GroupName { get; set; }
        public string Note { get; set; }
        public bool IsTotal { get; set; }
        public bool Marked { get; set; }
        public int AssetType { get; set; }
        public string HeadType { get; set; }
    }

    public class TempBalanceSheetModel
    {
        public string AccNo { get; set; }
        public string AccTitle { get; set; }
        public string SubAccOf { get; set; }
        public decimal OpeningBalance { get; set; }
        public decimal PrvAmount { get; set; }
        public decimal CurAmount { get; set; }
        public decimal LastAmount { get; set; }
        public string EntryId { get; set; }
        public string GroupName { get; set; }
        public string Notes { get; set; } // Mapping from SortNo
        public string SubGroupName { get; set; }
        public string SubEntryID { get; set; }
        public int IsDetail { get; set; }
    }

    public class PrintAttachmentsModel
    {
        public long EntryID { get; set; }
        public string Attachment_Name { get; set; }
        public string Attachment_Type { get; set; }
    }

    public class CashFlowMainHeadsModel
    {
        public long EntryID { get; set; }
        public string Description { get; set; }
        public string Notes { get; set; }
        public bool Marked { get; set; }
        public long CFHM_RefID { get; set; }
        public string CFMainHead { get; set; }
        public long CFH_RefID { get; set; }
        public string CF_HeadType { get; set; }
    }

    public class CashFlowType_DD_Model
    {
        public int EntryID { get; set; }
        public string CF_HeadType { get; set; }
    }

    public class CashFlowMain_DD_Model
    {
        public int EntryID { get; set; }
        public string CFMainHead { get; set; }
    }

    public class CashFlowReportHeads_Main_Model
    {
        public long EntryID { get; set; }
        public string Description { get; set; }
        public int SrNo { get; set; }
        public bool IsTotal { get; set; }
        public bool Marked { get; set; }
        public string NoteNo { get; set; }
        public List<CashFlowReportHeads_Detail_Model> DetailItems { get; set; } = new List<CashFlowReportHeads_Detail_Model>();
    }
    public class CashFlowReportHeads_Detail_Model
    {
        public long RefID { get; set; }
        public long Group_RefID { get; set; }
        public bool IsMain { get; set; } = false;
    }

    public class CashFlowReportHeadsDetailToLoadModel
    {
        public long EntryID { get; set; }
        public long RefID { get; set; }
        public string NoteNo { get; set; }
        public string Description { get; set; }
        public int SrNo { get; set; }
        public int Group_RefID { get; set; }
        public bool IsMain { get; set; }
        public string GroupName { get; set; }
        public string Note { get; set; }
        public bool IsTotal { get; set; }
        public bool Marked { get; set; }
    }
    public class CashFlowHeadsListModel
    {
        public int EntryID { get; set; }
        public string GroupName { get; set; }
    }

}