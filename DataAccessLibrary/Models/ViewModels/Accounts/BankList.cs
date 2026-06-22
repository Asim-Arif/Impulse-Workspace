using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        public bool ShowInCIP { get; set; }
        public bool COA { get; set; }
        public string AccNo { get; set; }
        public string AccType { get; set; }
        public string SubAccOf { get; set; }
        public string Remittance { get; set; }
        public string AccTitle { get; set; }
    }
}
