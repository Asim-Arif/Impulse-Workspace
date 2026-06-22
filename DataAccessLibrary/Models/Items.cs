//namespace DataAccessLibrary.Model

using Azure;

namespace DataAccessLibrary.Models
{
    public class Item
    {
        public string ItemID { get; set; }
        public string ItemName { get; set; }
        public string Unit { get; set; }
        public int Quantity { get; set; }
        public decimal Price { get; set; }
        public string Catagory { get; set; }
        public decimal StockQty { get; set; }
        public decimal Qty { get; set; }
        public decimal GQty { get; set; }
        public string BookInvoiceNo { get; set; }
        public DateTime BookDT { get; set; }
        public string BookDescription { get; set; }
        public string CustCode { get; set; }

        public string SBInvoiceNo { get; set; }
        public DateTime SBDT { get; set; }
        public string Customer { get; set; }
        public decimal TotalQty { get; set; }
        public decimal SoldQty { get; set; }
        public decimal BalQty { get; set; }
        public List<RMDetail> RMDetails { get; set; } = new List<RMDetail>();
    }

    public class RMDetail
    {
        public string RMCode { get; set; }
        public int RMQty { get; set; }
        public string ItemID { get; set; }
        public string ItemName { get; set; }
        public string Unit { get; set; }
        public decimal StockQty { get; set; }
        public decimal Qty { get; set; }
        public decimal GQty { get; set; }

    }
    public class Material
    {
        public Material() { }
        public Material(string materialCode,string materialName, string unit,int stockqty)
        {
            MaterialID=materialCode;
            MaterialName=materialName;
            Unit= unit;
            StockQty = stockqty;
        }
        public string MaterialID { get; set; }
        public string MaterialName { get; set; }
        public string Unit { get; set; }
        public int StockQty { get; set; }
    }

    public class SBList
    {
        public SBList() { }
        //public SBList(string materialCode, string materialName, string unit, int stockqty)
        //{
        //    MaterialID = materialCode;
        //    MaterialName = materialName;
        //    Unit = unit;
        //    StockQty = stockqty;
        //}
        public string SBInvoiceNo { get; set; }
        public DateTime SBDT { get; set; }
        public string Customer { get; set; }
        public decimal TotalQty { get; set; }
        public decimal SoldQty { get; set; }
        public decimal BalQty{ get; set; }
    }

}

