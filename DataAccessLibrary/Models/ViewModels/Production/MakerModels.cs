using System;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class MakerDto
    {
        public long VendID { get; set; }
        public string? VendID1 { get; set; }
        public string? VenderName { get; set; }
        public string? CompanyName { get; set; } // Label: Maker Name as per CNIC
        public string? Maker_Second_Name { get; set; } // Label: Maker Second Name
        public string? MakerType { get; set; }
        public string? AccNo { get; set; }
        public string? SubAccOf { get; set; }

        // Contact Info Panel
        public string? Phone1 { get; set; }
        public string? Phone2 { get; set; }
        public string? Phone3 { get; set; } // Fax
        public string? Mobile { get; set; }
        public string? NICNo { get; set; }
        public string? NTNNo { get; set; }
        public string? Address { get; set; }
        public string? BankAccNo { get; set; }
        public int? MaxLimit { get; set; } // Legacy DataField="Fax1"
        public int? TimeLimit { get; set; } // Legacy DataField="Fax2"

        // Tab 0: General Info
        public bool AuthRequired { get; set; }
        public bool ShowRateOnPO { get; set; }
        public string? MakerNameUrdu { get; set; }
        public string? CNIC_PDF_FileName { get; set; }

        // Tab 2: Other Info
        public string? PaymentTerms { get; set; }
        public string? RefBy { get; set; }
        public string? RefByFName { get; set; }
        public string? RefByPhone1 { get; set; }
        public string? RefByPhone2 { get; set; }
        public string? RefByAddress { get; set; }
        public float? ExcessQtyPercentage { get; set; }
        public byte? MaximumRcvingsAgainstPO { get; set; }
        public int? MakerCapacity { get; set; }

        public bool Active { get; set; } = true;
    }

    public class ProcessModel
    {
        public int ProcessID { get; set; }
        public string Description { get; set; } = string.Empty;
        public string Code { get; set; } = string.Empty;
        public int SNo { get; set; }
        public bool IsSelected { get; set; }
    }

    public class MakerCategoryModel
    {
        public string AccNo { get; set; } = string.Empty;
        public string AccTitle { get; set; } = string.Empty;
    }
}
