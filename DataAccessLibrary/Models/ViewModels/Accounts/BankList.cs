using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Accounts
{
    public class BankListModel
    {
        public int BankID { get; set; }
        public string Bank { get; set; }
        public string Branch { get; set; }
        public string City { get; set; }
        public string Phone1 { get; set; }
        public string Phone2 { get; set; }
        public string FaxNo { get; set; }
        public string Email { get; set; }
        public string Manager { get; set; }
        public string Address { get; set; }
        public string BankAccNo { get; set; }
        public string IBAN { get; set; }
        public string SwiftCode { get; set; }
        public string AccNo { get; set; }
        public string AccType { get; set; }
        public string SubAccOf { get; set; }
        public string Remittance { get; set; }
        public string AccTitle { get; set; }
    }

    public class BankBranchRecord
    {
        public int BankID { get; set; }
        public string Bank { get; set; } = string.Empty;
        public string Branch { get; set; } = string.Empty;
        public string City { get; set; } = string.Empty;
        public string Phone1 { get; set; } = string.Empty;
        public string Phone2 { get; set; } = string.Empty;
        public string Fax { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public string Manager { get; set; } = string.Empty;
        public string Address { get; set; } = string.Empty;
        public string AccNo { get; set; } = string.Empty;
        public string Remittance { get; set; } = string.Empty;
        public int? CompanyRefID { get; set; }
        public int AccountsCount { get; set; }
    }

    public class BankAccountItem
    {
        public int BankID { get; set; }
        public string Bank { get; set; } = string.Empty;
        public string Branch { get; set; } = string.Empty;
        public string City { get; set; } = string.Empty;
        public string AccNo { get; set; } = string.Empty;
        public string AccTitle { get; set; } = string.Empty;
        public string BankAccNo { get; set; } = string.Empty;
        public string Type { get; set; } = "Current";
        public bool DefaultBank { get; set; }
        public bool AllowOD { get; set; }
        public decimal? ODAmt { get; set; }
        public float? ODRate { get; set; }
        public string Chqformat { get; set; } = string.Empty;
        public decimal Balance { get; set; }
        public decimal OpenBal { get; set; }
        public DateTime? OpenDate { get; set; }
        public bool Active { get; set; } = true;
        public int ChqBooksCount { get; set; }
    }

    public class SaveBankBranchDto
    {
        public int BankID { get; set; }
        public string Bank { get; set; } = string.Empty;
        public string Branch { get; set; } = string.Empty;
        public string City { get; set; } = string.Empty;
        public string Phone1 { get; set; } = string.Empty;
        public string Phone2 { get; set; } = string.Empty;
        public string Fax { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public string Manager { get; set; } = string.Empty;
        public string Address { get; set; } = string.Empty;
        public string Remittance { get; set; } = string.Empty;
        public int? CompanyRefID { get; set; }
    }

    public class SaveBankAccountDto
    {
        public int BankID { get; set; }
        public string AccNo { get; set; } = string.Empty;
        public string AccTitle { get; set; } = string.Empty;
        public string BankAccNo { get; set; } = string.Empty;
        public string Type { get; set; } = "Current";
        public bool DefaultBank { get; set; }
        public bool AllowOD { get; set; }
        public decimal? ODAmt { get; set; }
        public float? ODRate { get; set; }
        public string Chqformat { get; set; } = string.Empty;
        public decimal OpenBal { get; set; }
        public DateTime? OpenDate { get; set; }
        public bool IsNew { get; set; }
    }

    public class OverdraftConfigDto
    {
        public string AccNo { get; set; } = string.Empty;
        public bool AllowOD { get; set; }
        public decimal? ODAmt { get; set; }
        public float? ODRate { get; set; }
    }

    public class BankOperationResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public string GeneratedAccNo { get; set; } = string.Empty;
        public int BankID { get; set; }
    }
}
