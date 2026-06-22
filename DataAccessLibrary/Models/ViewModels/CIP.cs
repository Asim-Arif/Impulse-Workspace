using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccessLibrary.Models
{
    public class CipLoungeEntry
    {
        public DateTime CIPDT { get; set; }
        public string CIPInvoiceNo { get; set; }
        public string CipSourceType { get; set; }
        public string PassengerName { get; set; }
        public string SeatNo { get; set; }
        public string FlightNo { get; set; }
        public string CardNo { get; set; }
        public string Remarks { get; set; }
        public string authcode { get; set; }
        
        public decimal Amount { get; set; }
        public decimal TotalAmount { get; set; }
        public decimal RcvdAmount { get; set; }
        public decimal BalanceAmount { get; set; }
        public int Totalperson { get; set; }
        public string CIPCardNo { get; set; }
        public string CipRemarks { get; set; }
        public string CipSource { get; set; }
        public string CipType { get; set; }
        public int CipPersons { get; set; }
        public decimal CipNetAmount { get; set; }
        public decimal TotalRcvdAmount { get; set; }
        public string AttType { get; set; }
        public int Dir_RefID { get; set; }
        public string DirectorName { get; set; }
        public int BankID { get; set; }
        public string BankName { get; set; }
        public string EmpID { get; set; }
        public string EmpName { get; set; }
        public string Code { get; set; }
        public string AirlineName { get; set; }
        public string usernameforinvoices { get; set; }
        public string CipSourceTypesForFilter { get; set; }
        public string CipSourceTypesForFilter_values { get; set; }
        public int EntryID { get; set; }
        public DateTime CIPDTFrom { get; set; }
        public DateTime CIPDTTo { get; set; }
        public int LD_RefID { get; set; }

        public int CIP_Approval_RefID { get; set; }
        public string CIP_Approval_ForwardBy { get; set; }
        public string CIP_Approval_ForwardTo { get; set; }

        public int FolioNo { get; set; }
        public string CnicNo { get; set;}
        public string machineIP { get; set; }
        public Boolean bCOA { get; set; }
        public string AuthCardNo { get; set; }

        public List<CIPDetail> CIPDetails { get; set; } = new List<CIPDetail>();
        public List<CIPDetail_Typewise> CIPDetails_tw { get; set; } = new List<CIPDetail_Typewise>();
        
    }

    public class CIPDetail
    {
        public string PassengerName { get; set; }
        public string SeatNo { get; set; }
        public string FlightNo { get; set; }
        public string CardNo { get; set; }
        public string authcode { get; set; }
        
        public decimal Amount { get; set; }
        public int LD_RefID { get; set; }
        public string AuthCardNo { get; set; }
    }

    public class CIPList
    {
        //public CIPList() { }
        public string CIPInvoiceNo { get; set; }
        public DateTime CIPDT { get; set; }
        public string CipSource { get; set; }
        public string CipType { get; set; }
        public int CipPersons { get; set; }
        public decimal CipNetAmount { get; set; }
        public decimal TotalRcvdAmount { get; set; }
        public string CipSourceType { get; set; }
        public string AttachmentType { get; set; }
        public string PassengerName { get; set; }
        public string SeatNo { get; set; }
        public string FlightNo { get; set; }
        public string authcode { get; set; }
        
        public int Amount { get; set; }
        public int EntryID { get; set; }
    }

    public class CIPDetail_Typewise
    {
        public int LD_RefID { get; set; }
    }

    public class CIPBanks
    {
        public string bankname { get; set; }
        public int bankid { get; set; }
        public Boolean bCOA { get; set; }
    }

    public class ExcelItem
    {
        public string DirectorName { get; set; }
        public int FolioNo{ get; set; }
        public string CnicNo { get; set; }
    }

    public class UserSessionService
    {
        public string UserName { get; set; }
    }

}


