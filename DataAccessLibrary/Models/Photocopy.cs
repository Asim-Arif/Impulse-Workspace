using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace DataAccessLibrary.Models
{
    public class CopyEntry
    {
        public DateTime PCDT { get; set; }
        public string CopyInvoiceNo { get; set; }
        public string CopyType { get; set; }
        public decimal Qty { get; set; }
        public decimal Rate { get; set; }
        public decimal Amount { get; set; }
        public string Remarks { get; set; }
        public string PhotoCopyType { get; set; }
        public int EntryID { get; set; }
        public string ForwardBy { get; set; }
        public string ForwardTo { get; set; }

        public DateTime DTFrom { get; set; }
        public DateTime DTTo { get; set; }
        public string AttType { get; set; }

        public List<CopiesDetail> CIPDetails { get; set; } = new List<CopiesDetail>();
    }

    public class CopiesDetail
    {
        public DateTime PCDT { get; set; }
        public string CopyInvoiceNo { get; set; }
        public string CopyType { get; set; }
        public decimal Qty { get; set; }
        public decimal Rate { get; set; }
        public decimal Amount { get; set; }
        public string Remarks { get; set; }
        public int EntryID { get; set; }

    }


    public class CopiesList
    {
        public DateTime PCDT { get; set; }
        public string CopyInvoiceNo { get; set; }
        public string CopyType { get; set; }
        public decimal Qty { get; set; }
        public decimal Rate { get; set; }
        public decimal Amount { get; set; }
        public string Remarks { get; set; }
        public int EntryID { get; set; }
    }

    //public class CopiesDetail
    //{
    //    public int LD_RefID { get; set; }
    //}

}



