using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;

namespace DataAccessLibrary.Models.ViewModels.Accounts
{
    public class Online_PVModelView
    {
        public int EntryID { get; set; }   // Identity PK

        public DateTime DT { get; set; }

        public string? SIAL_Bank_Acc_No { get; set; }
        public string? Debit_Acc_No { get; set; }
        public string? Tax_Acc_No { get; set; }

        public string? UserName { get; set; }
        public string? MachineName { get; set; }

        public DateTime? DTEntry { get; set; }

        public byte? Status { get; set; }

        public string? UserName_Approved { get; set; }
        public string? MachineName_Approved { get; set; }
        public DateTime? DTEntry_Approved { get; set; }

        public string? VchrNo { get; set; }

        public int? Transaction_Type { get; set; }

        public string? Remarks { get; set; }

        public bool? SMS_Sent { get; set; }
        public string? SMS_Request { get; set; }
        public string? SMS_Response { get; set; }

        public string? SMS_UserName { get; set; }
        public string? SMS_MachineName { get; set; }
        public DateTime? DTEntry_SMS { get; set; }

        public int? BOD_Meetings_RefID { get; set; }

        public bool? Email_Sent { get; set; }
        public string? Email_UserName { get; set; }
        public string? Email_MachineName { get; set; }
        public DateTime? DTEntry_Email { get; set; }

        public List<Online_PV_DetailsModelView> LineItems { get; set; } = new List<Online_PV_DetailsModelView>();
        public int? TotalDebit => LineItems.Sum(li => li.Debit);
        public int? TotalTax => LineItems.Sum(li => li.Tax);
    }
    public class Online_PV_DetailsModelView 
    {
        public int EntryID { get; set; }   // Identity PK

        public int? OL_PV_RefID { get; set; }   //Reference Key
        public int? SD_RefID { get; set; }  //Sial Directors Reference Key

        public string? Bank_Acc_No { get; set; }
        public string? Bank_Code { get; set; }
        public string? Bank_Acc_Title { get; set; }
        public string? Bank_Name { get; set; }

        public int? Debit { get; set; }
        public int? Tax { get; set; }

        public byte? Status { get; set; }

        public string? Response_Raw { get; set; }
        public string? Response_Code { get; set; }
        public string? Response_Description { get; set; }
        public string? Response_UnqiueTxnNo { get; set; }

        public string? Request_Raw { get; set; }

        public string? SMS_Status { get; set; }
        public string? SMS_Response { get; set; }

        public string? ChqNo { get; set; }
        public string? Advising { get; set; }

        public string? UserName_Post { get; set; }
        public string? MachineName_Post { get; set; }
        public DateTime? DTEntry_Post { get; set; }

        public string? Email_SentTo { get; set; }
    }

    public class SIALDirectorViewModel
    {
        public int EntryID { get; set; }

        public string? DirectorName { get; set; }
        public string? NICNo { get; set; }
        public string? FatherName { get; set; }
        public string? Address { get; set; }
        public string? Designation { get; set; }
        public string? Nationality { get; set; }
        public string? NationalityOrigin { get; set; }
        public string? OtherBusiness { get; set; }

        public DateTime? DTPresentAppointment { get; set; }
        public string? Changes { get; set; }

        public int? SM_RefID { get; set; }

        public byte[]? Pic { get; set; }   // SQL image → byte[]

        public string? Phone { get; set; }
        public string? Email { get; set; }
        public string? Fax { get; set; }
        public string? Cell { get; set; }

        public bool? InternalReportingOnly { get; set; }

        public string? Phone2 { get; set; }
        public string? PhoneResidence { get; set; }

        public string? Company { get; set; }
        public string? Nominee { get; set; }

        public int? FolioNo { get; set; }

        public string? Bank_Account_No { get; set; }
        public string? Bank_Account_Title { get; set; }
        public string? Bank_Name { get; set; }
        public string? Bank_Code { get; set; }
        public string? Bank_Contact_No { get; set; }

        public string? ALD_FolioNo { get; set; }

        public DateTime? DTentry { get; set; }

        public int? Active { get; set; }
        public int? SD_RefID {  get; set; } //to handle Processed for BOD Metting Payment.
    }
    public class SIALDirectorPaymentViewModel
    {
        // The core director data (You can also just inherit from SIALDirectors)
        public SIALDirectorViewModel Director { get; set; } = new();

        // UI and Logic Fields
        public bool IsChecked { get; set; }
        public bool IsHighlighted { get; set; }
        public bool IsOnlineTitleFetched { get; set; }
        public bool IsProcessed { get; set; }
        public int Debit { get; set; }
        public int Tax { get; set; }

        // Calculated property for convenience
        public int Net => Debit - Tax;
    }
    public class BOD_Meeting_Payment_List_ViewModel 
    {
        public int EntryID { get; set; }
        public int SNo { get; set; }
        public DateTime DT { get; set; }
        public string SIAL_Bank_Acc_No { get; set; }
        public string SIAL_Bank_Acc_Title { get; set; }
        public string Debit_Acc_No { get; set; }
        public string Debit_Acc_Title { get; set; }
        public string Tax_Acc_No { get; set; }
        public string Tax_Acc_Title { get; set; }
        public int BankID { get; set; }
        public int TotalDebit { get; set; }
        public int TotalTax { get; set; }
        public int TotalEntries { get; set; }
        public int Status { get; set; }
        public string Remarks { get; set; }
        public string PostedVchrNo { get; set; }
        public int Transaction_Type { get; set; }
        public int TotalNetAmount => TotalDebit - TotalTax;
        public string strStatus => (Status == 0 && !string.IsNullOrEmpty(PostedVchrNo)) ? "Voucher Posted" :
                               (Status == 1 && !string.IsNullOrEmpty(PostedVchrNo)) ? "Approved & Posted" :
                               "Under Process";

        public string strTransaction_Type => Transaction_Type switch
        {
            0 => "Funds Transfer",
            1 => "Corporate Cheque",
            2 => "Ebiz Plus",
            3 => "Standard Corporate Cheque",
            4 => "Real Time Gross Settlement",
            5 => "Cash over Counter",
            _ => "Unknown" // Default case
        };

        // Returns a Bootstrap class for the <tr> tag
        public string RowClass => (Status == 0 || Status == 1) && !string.IsNullOrEmpty(PostedVchrNo)
                                  ? "table-success-light" // We will define this custom light green
                                  : "table-warning-light"; // We will define this custom light orange

        // Returns a Bootstrap class for a Badge/Pill
        public string BadgeClass => (Status == 0 || Status == 1) && !string.IsNullOrEmpty(PostedVchrNo)
                                    ? "bg-success"
                                    : "bg-warning text-dark";
    }
    public class BOD_Meeting_Payment_Detail_ViewModel
    {
        public int EntryID { get; set; }
        public int OL_PV_RefID { get; set; }
        public int SD_RefID { get; set; }
        public string Bank_Acc_No { get; set; }
        public string Bank_Code { get; set; }
        public string Bank_Acc_Title { get; set; }
        public string Bank_Name { get; set; }
        public int Debit { get; set; }
        public int Tax { get; set; }
        public int Status { get; set; }
        public string SMS_Status { get; set; }
        public int FolioNo { get; set; }
        public string DirectorName { get; set; }
        public int BOD_Meetings_RefID { get; set; }
        public string ChqNo { get; set; }
        public string Advising {  get; set; }
        public int Transaction_Type { get; set; }
        public int BankID { get; set; }
        public string Bank_Accurate_Name { get; set; }
        public string Bank_Code_Name { get; set; }
        public string VchrNo { get; set; }

    }
    public class Online_Payment_Voucher_List_ViewModel
    {
        public string VchrNo {  set; get; }
        public DateTime VDate { get; set; }
        public string Payee { set; get; }
        public int Amount { get; set; }
        public string ChequeNo { set; get; }
        public string Advising { set; get; }
        public bool Online_Posting { set; get; }
        public int Online_Vchr_Type { set; get; }
        public string M1No { set; get; }
        public string Bank { set; get; }

        public string Bank_Code { get; set; }
        public string Branch { set; get; }
        public string City { set; get; }
        public string AccNo { set; get; }
        public int BankID {  set; get; }

        public bool IsChecked { set; get; }
        public string strAccType => (BankID==60 || BankID==86) ? "Saving A/C" : "Current A/C";
        public string strStatus => (Online_Posting) ? "Posted" : "Pending";
        
        public string strTransaction_Type => Online_Vchr_Type switch
        {
            2 => "Corporate Cheque",
            1 => "Funds Transfer",            
            _ => "Unknown" // Default case
        };
        public string RowClass => (Online_Posting)
            ? "table-success-light" // We will define this custom light green
            : "table-warning-light"; // We will define this custom light orange
        public string BadgeClass => (Online_Posting)
            ? "bg-success"
            : "bg-warning text-dark";

    }
    public class Trial_Balance_ViewModel
    {
        public string AccNo { set; get; }
        public string AccTitle { set; get; }
        public decimal Opening_Debit { get; set; }
        public decimal Opening_Credit { get; set; }
        public decimal Debit { get; set; }
        public decimal Credit { get; set; }
        public decimal Balance_Debit { get; set; }
        public decimal Balance_Credit { get; set; }
        public int EntryType { get; set; }


        public string RowClass => EntryType switch
        {
            0 => "table-success-light", // We will define this custom light green
            1 => "table-warning-light", // We will define this custom light orange
            2 => "table-primary-light", // We will define this custom light orange
            _ => ""
        };
        
    }
    public class Bank_Balance_Statement_ViewModel
    {
        public string AccNo { set; get; }
        public string Bank { set; get; }
        public string Branch { set; get; }
        public string BankAccNo { set; get; }
        public decimal Balance { get; set; }
        public string ChqFormat { set; get; }

        public string RowClass => (Balance < 0) ? "table-warning-light" : "table-success-light";        

    }
    public class Cheque_Receiving_ViewModel 
    {
        public int EntryID { set; get; }
        public DateTime DT { set; get; }
        public string PartyName { set; get; }
        public string ChqNo { set; get; }
        public string DepositBank { set; get; }


        //[Range(1, int.MaxValue, ErrorMessage = "Please select a valid Employee")] // Use this if EmpID is a numeric string
        [Required(ErrorMessage = "Please select an Employee")]
        public string EmpID { set; get; }

        [Range(0.01, double.MaxValue, ErrorMessage = "Amount must be greater than zero")]
        public decimal Amount { set; get; }
        public string ReceivedFrom { set; get; }
        public DateTime DTChq { set; get; }
        public string SlipNo { set; get; }
        public DateTime DTSlip { set; get; }
        public string Name { set; get; }
        public string DeptName { set; get; }

    }
    public class Important_Tasks_ViewModel
    {
        public int EntryID { set; get; }
        public DateTime DT { set; get; }
        public string Title { set; get; }
        public string Particulars { set; get; }
        public string DutyAssignedTo { set; get; }
        public string Status { set; get; }
        public DateTime DTDue { set; get; }
        public bool OK { set; get; }
        
    }
    public class Bills_Received_ViewModel
    {
        public int EntryID { set; get; }
        public int RefNo { set; get; }
        public string Description { set; get; }
        public int Amount { set; get; }
        public DateTime DTIn { set; get; }
        public DateTime? DTDue { set; get; }
        public string ReceivedFrom { set; get; }
        public string ReceivedBy { set; get; }
        public string Objection { set; get; }
        public string HandedOverTo { set; get; }
        public DateTime DTHandedOver { set; get; }
        public string Status { set; get; }
        public string Remarks { set; get; }
        public string Party { set; get; }
        public DateTime DTIn_DateOnly { set; get; }

    }
    public class NTN_CNIC_ViewModel
    {
        public int EntryID { set; get; }
        
        public string ReferenceNo { set; get; }        
        public string Name { set; get; }
        public string NTNNo { set; get; }
        public string CNICNo { set; get; }

    }
    public class Letter_Dispatch_Register_ViewModel
    {
        public int EntryID { set; get; }
        public DateTime DT { set; get; }
        public string RefNo { set; get; }
        public string PartyName { set; get; }
        public string Address { set; get; }
        public string CellNo { set; get; }
        public string CourierName { set; get; }
        public string CourierNo { set; get; }

    }

}