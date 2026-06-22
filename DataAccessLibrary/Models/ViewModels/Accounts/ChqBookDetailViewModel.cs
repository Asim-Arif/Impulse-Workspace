using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;


namespace DataAccessLibrary.Models.ViewModels.Accounts
{
    public class ChqBookDetailViewModel
    {
        public long ChqBookNo { get; set; }
        public string AccNo { get; set; }
        public string StartingFrom { get; set; }
        public int Chqs { get; set; }
        public string ManualNo { get; set; }
        public string ChqBookDetail { get; set; }
        public int Chq_Type { get; set; }
        public string UserName { get; set; }
        public string MachineName { get; set; }
        public DateTime DT { get; set; }
        public int ChqsLeft { get; set; }

    }

    public class ChqListModel
    { 
        public long ChqBookNo { get; set; }
        public string ChqNo { get; set; }
        public string Chqtype { get; set; }
        public DateTime CDate { get; set; }
        public DateTime ChequeDate { get; set; }
        public string AccTitle { get; set; }
        public string AccNo { get; set; }
        public string Description { get; set; }
        public decimal Amount { get; set; }
        public string Payee { get; set; }
        public string ClearanceDT { get; set; }
        public bool Bounced { get; set; }
        public bool Issued { get; set; }
        public string DeleteDescription { get; set; }
        public string ChqStatus { get; set; }
        public double CSNo { get; set; }
        public string CancelReason { get; set; }
        public string chqdesignation { get; set; }
        public string chqcompany { get; set; }
        public string chqpayee { get; set; }
        public DateTime chqprintdate { get; set; }
        public string UserName { get; set; }
        public string MachineName { get; set; }

    }
}