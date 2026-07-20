using System;

namespace DataAccessLibrary.Models.ViewModels.Accounts
{
    public class ChequeDetailModel
    {
        public DateTime CDate { get; set; }
        public string ChequeNo { get; set; }
        public string Bank { get; set; }
        public string Branch { get; set; }
        public string Description { get; set; }
        public decimal Amount { get; set; }
        public string ChequeType { get; set; }
        public DateTime? ChequeDate { get; set; }
        public bool Posted { get; set; }
        public bool Bounced { get; set; }
        public DateTime? ClearanceDT { get; set; }
        public bool Payment { get; set; }
    }
}
