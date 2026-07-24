using System;
using System.ComponentModel.DataAnnotations;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class NewVendorViewModel
    {
        public int VendID { get; set; }

        public string AccNo { get; set; } = string.Empty;

        [Required(ErrorMessage = "Vendor Name is required")]
        public string VendorName { get; set; } = string.Empty; // Maps to AccTitle in Accounts

        [Required(ErrorMessage = "Please select a Category")]
        public string ParentAccNo { get; set; } = string.Empty; // Category from Accounts where Parent=1

        public decimal OpeningBalance { get; set; }

        // "DR" or "CR"
        public string OpeningBalanceType { get; set; } = "DR";

        public decimal CurrentBalance { get; set; }

        public bool Inactive { get; set; }

        public string Address { get; set; } = string.Empty;

        public string ContactPerson { get; set; } = string.Empty;

        public string Phone1 { get; set; } = string.Empty;
        public string Phone2 { get; set; } = string.Empty;
        public string Mobile { get; set; } = string.Empty;
        public string CPhone { get; set; } = string.Empty;
        public string CEmail { get; set; } = string.Empty;

        public string Fax1 { get; set; } = string.Empty;
        public string Fax2 { get; set; } = string.Empty;

        public string MakerNo { get; set; } = string.Empty;
        public string BankAccNo { get; set; } = string.Empty;

        public int ProcessID { get; set; }

        public bool ImportVendor { get; set; }

        public string VendorDescription { get; set; } = string.Empty;

        public string VendorNameUrdu { get; set; } = string.Empty;

        // Base64 representations for pictures
        public string VendorPic { get; set; } = string.Empty;
        public string VendorSig { get; set; } = string.Empty;
    }
}
