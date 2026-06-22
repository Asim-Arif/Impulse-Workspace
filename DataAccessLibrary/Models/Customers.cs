using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccessLibrary.Models
{
    public class Customers
    {
        public Customers() { }
        public Customers(string custcode, string customername)
        {
            CustCode = custcode;
            CustomerName = customername;
        }
        public string CustCode { get; set; }
        public string CustomerName { get; set; }

        public List<CustDetail> CustDetails { get; set; } = new List<CustDetail>();
    }

    public class CustDetail
    {
        public string CustCode { get; set; }
        public string CustomerName { get; set; }
    }

    public class SaleBooking
    {
        //public string BookInvoiceNo { get; set; }
        ////public string BookDT { get; set; }
        //public string BookDescription { get; set; }
    }
}
