using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    public class ShortTermSheetItemViewModel
    {
        public string EmpID { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string EmpIDOld { get; set; } = string.Empty;
        public string Designation { get; set; } = string.Empty;
        public string DeptName { get; set; } = string.Empty;
        public string DeptID { get; set; } = string.Empty;
        public decimal StartingSalary { get; set; }
        public double PresentDays { get; set; }
        public double Leaves { get; set; }
        public double WPLeaves { get; set; }
        public decimal Payed { get; set; }
        public decimal AmtToDeduct { get; set; }
        public double OverTimeHours { get; set; }
        public double OTRate { get; set; } = 1.0;
        public decimal CalculatedPayable { get; set; }
        public decimal Demand { get; set; }
        public decimal Approved { get; set; }
    }

    public class ShortTermSheetSaveDto
    {
        public DateTime Date { get; set; }
        public string EmpID { get; set; } = string.Empty;
        public decimal Demand { get; set; }
        public decimal Approved { get; set; }
    }

    public enum ShortTermPostingMode
    {
        Cash = 0,
        Bank = 1,
        OtherAccount = 2
    }

    public class ShortTermPostingDto
    {
        public DateTime Date { get; set; }
        public string DeptID { get; set; } = "0";
        public ShortTermPostingMode PaymentMode { get; set; } = ShortTermPostingMode.Cash;
        public string PaymentAccNo { get; set; } = string.Empty;
        public string BankAccNo { get; set; } = string.Empty;
        public string BankId { get; set; } = string.Empty;
        public string ChequeNo { get; set; } = string.Empty;
        public string ChequeType { get; set; } = "Open";
        public DateTime ChequeDate { get; set; } = DateTime.Today;

        public bool DontPostVoucher { get; set; } = false;
        public List<ShortTermSheetItemViewModel> Items { get; set; } = new();
    }

    public class ShortTermPostingStatusModel
    {
        public bool IsPosted { get; set; }
        public string CashVchrNo { get; set; } = string.Empty;
        public string BankVchrNo { get; set; } = string.Empty;
        public string PostedBy { get; set; } = string.Empty;
        public DateTime? PostedDate { get; set; }
    }
}

