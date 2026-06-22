using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;


namespace DataAccessLibrary.Models.ViewModels.Accounts
{
    public class Vouchers_MainViewModel
    {
        public int EntryID { get; set; }

        public DateTime DT { get; set; }

        public string? VchrNo { get; set; }

        public string? Remarks { get; set; }

        public DateTime? DueDate { get; set; }

        public bool? AdjVchr { get; set; }

        public string? SlipNo { get; set; }

        public string? ReceivedFrom { get; set; }

        public int? Approved { get; set; }

        public DateTime? DTEntry { get; set; }

        public string? UserName { get; set; }

        public string? MachineName { get; set; }

        public string? VchrNo_App { get; set; }

        public string? Attachment_Type { get; set; }

        public string? PDFAttachments { get; set; }

        public int? BankID { get; set; }

        public string? ChqType { get; set; }

        public DateTime? ChqDate { get; set; }

        public string? ChqNo { get; set; }

        public string? BankAccNo { get; set; }

        public string? OrderNo { get; set; }

        public int? CheckedByAudit { get; set; }

        public int? LTVchr { get; set; }

        public DateTime? DTSalary { get; set; }

        public string? Remarks_Audit { get; set; }

        public int? TaxType { get; set; }

        public string? HeadType { get; set; }

        public float? TaxAmount { get; set; }

        public float? TaxRate { get; set; }

        public string? Reason { get; set; }

        public bool? CaptilizeVchr { get; set; }

        public string ForwardTo { get; set; }
        public List<VouchersApprovalDetailViewModel> LineItems { get; set; } = new List<VouchersApprovalDetailViewModel>();

        // Calculated properties
        public decimal TotalDebit => LineItems.Sum(li => li.Debit);
        public decimal TotalCredit => LineItems.Sum(li => li.Credit);

    }
    public class VouchersApprovalViewModel
    {
        public int EntryID { get; set; }

        public int? RefID { get; set; }

        public string? ForwardBy { get; set; }

        public DateTime? DTEntry { get; set; }

        public string? MachineName { get; set; }

        public string? ForwardTo { get; set; }

        public DateTime? DTEntry_Done { get; set; }

        public string? MachineName_Done { get; set; }

        public string? Remarks { get; set; }

        public bool? Approval { get; set; }
    }

    
    public class VouchersApprovalDetailViewModel
    {
        public int EntryID { get; set; }

        public int? RefID { get; set; }

        public string? VchrNo { get; set; }

        public string? AccNo { get; set; }
        public string? AccTitle { get; set; }
        public string? Description { get; set; }

        public decimal Debit { get; set; }

        public decimal Credit { get; set; }

        public decimal? Balance { get; set; }

        public string? DpstSlip { get; set; }

        public int? CSNo { get; set; }

        public int? CS_RefID { get; set; }
        public string? CS_Description { get; set; }
        public bool IsSystemLine { get; set; } = false; //this is for UI handling, it's not matched to database field.

        // --- NEW PROPERTY FOR SUB-DETAILS ---
        public List<VAD_Voucher_DetailViewModel> AdjustmentDetails { get; set; } = new();

        // Helper property to see if the row has adjustments
        public bool HasAdjustments => AdjustmentDetails.Any();

        // Helper to calculate total adjusted amount for this line
        public double TotalAdjustedAmount => AdjustmentDetails.Sum(x => x.AdjAmount ?? 0);
    }
    public class VAD_Voucher_DetailViewModel
    {
        public int EntryID { get; set; }              // IDENTITY, PK

        public int? RefID { get; set; }                // FK → Vouchers_Approval_Detail.EntryID

        public string? VchrNo { get; set; }            // varchar(50)

        public double? Sno { get; set; }               // float in SQL → double in C#

        public string? AccNo { get; set; }             // varchar(50)

        public float? Amount { get; set; }             // real

        public float? BalanceAmount { get; set; }      // real

        public float? AdjAmount { get; set; }          // real

        public float? RemainAmount { get; set; }       // real

        public float? VchrAmount { get; set; }         // real

        
    }
    public class Voucher_MainSaveResult
    {
        public string VchrNo { get; set; }
        public int EntryID { get; set; }
    }

}
