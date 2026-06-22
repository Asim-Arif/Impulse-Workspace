using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;


namespace DataAccessLibrary.Models.ViewModels.Accounts
{
    public class VoucherViewModel 
    {
        public string VchrNo { get; set; }
        public string UserName { get; set; }
        public string MachineName { get; set; }
        public DateTime DT { get; set; }
        public string ApprovedBy { get; set; }
        public bool PostedThroughJVForm { get; set; }
        public string Notes { get; set; }
        public DateTime? DueDate { get; set; }
        public string Payee { get; set; }
        public byte JVType { get; set; }
        public bool VchrEdited { get; set; }
        public string Edit_Del_VchrNo { get; set; }
        public string Edit_Reason { get; set; }
        public string UserName_Original { get; set; }
        public string MachineName_Original { get; set; }
        public string UserName_Approved { get; set; }
        public string MachineName_Approved { get; set; }
        public DateTime DTEntry_Approved { get; set; }
        public string Response_Raw { get; set; }
        public string Response_Code { get; set; }
        public string Response_Description { get; set; }
        public string Response_UnqiueTxnNo { get; set; }
        public string Request_Raw { get; set; }
        public bool Online_Posting { get; set; }
        public bool Online_Vchr { get; set; }
        public int Online_Vchr_Type { get; set; }
        public string Advising {  get; set; }
        public string UserName_Post { get; set; }
        public string MachineName_Post { get; set; }
        public DateTime DTEntry_Post { get; set; }
        public string M1No { get; set; }
        public DateTime M1DT { get; set; }
        public DateTime DT_Original { get; set; }
        public DateTime? BillDate { get; set; }
        public string BillNo { get; set; }
        public string OrderNo { get; set; }
        public int TaxType { get; set; }
        public string HeadType { get; set; }
        public float TaxAmount { get; set; }
        public float TaxRate { get; set; }
        public string Reason { get; set; }
        public bool CaptilizeVchr { get; set; }
        public int VendID { get; set; }
        public string EmpID { get; set; }
        public int PaymentToEmp {  get; set; }


        public List<VoucherLineItemViewModel> LineItems { get; set; } = new List<VoucherLineItemViewModel>();

        // --- New Optional Cheque Property ---
        // If the voucher is a Bank Payment Voucher (BPV), this will be populated.
        // Otherwise, it will be null.
        public ChequeViewModel? ChequeDetails { get; set; } // The '?' denotes nullable reference type (C# 8.0+)

        // --- Utility Property to check if Cheque is involved ---
        public bool IsChequePayment => ChequeDetails != null;

        // Calculated properties
        public decimal TotalDebit => LineItems.Sum(li => li.Debit);
        public decimal TotalCredit => LineItems.Sum(li => li.Credit);

    }
    public class VoucherLineItemViewModel
    {
        public double SNo { get; set; }
        public DateTime VDate { get; set; }
        public string VchrNo { get; set; }
        public string AccNo { get; set; }
        public string AccTitle { get; set; }
        public string Description { get; set; }        
        public decimal Debit { get; set; }
        public decimal Credit { get; set; }
        public decimal Balance { get; set; }
        public string DpstSlip { get; set; }
        public long CSNo { get; set; }
        public int CS_RefID { get; set; }
        public string CS_Description { get; set; }

        public bool IsSystemLine { get; set; } = false; //this is for UI handling, it's not matched to database field.

    }
    public class ChequeViewModel
    {
        // Primary Key or Foreign Key linking to VoucherInfo (if needed)
        public long SNo { get; set; }

        [DataType(DataType.Date)]
        public DateTime CDate { get; set; }

        public string AccNo { get; set; }
        public string Description { get; set; }
        public decimal Amount { get; set; }
        public int BankID { get; set; }
        public string ChequeNo { get; set; }
        public int ChqBookNo { get; set; }
        public string chequeType { get; set; }

        [DataType(DataType.Date)]
        public DateTime ChequeDate { get; set; }
        
        public bool Posted { get; set; }
        public bool Payment { get; set; }
        public bool Bounced { get; set; }
        public bool ChqIsDue { get; set; }

        public string Payee { get; set; }
        public string BankAccNo { get; set; }
        public DateTime? ClearanceDT { get; set; }
        public bool ChqPrintingDone { get; set; }
        public double? Vouchers_SNo { get; set; }

    }

    public class BankAccountInfo
    {
        public int BankID { get; set; }
        public string AccNo { get; set; }
        public string AccTitle { get; set; }
        public string Bank { get; set; }
        public string Branch { get; set; }
        public string City { get; set; }
        public string BankAccNo { get; set; }
        public bool DefaultBank { get; set; }
        public bool AllowOD { get; set; }
        public double ODRate { get; set; }
        public long ODAmt { get; set; }
    }

    public class ChequeInfo
    {
        public int ChqBookNo { get; set; }
        public string ChqNo { get; set; }
        public bool Issued { get; set; }
    }

}
