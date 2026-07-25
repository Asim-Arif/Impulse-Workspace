using System;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class StockLedgerTransactionModel
    {
        public string TransactionID { get; set; }
        public DateTime TransactionDate { get; set; }
        public string Description { get; set; }
        public string Reference { get; set; }
        public double Quantity { get; set; }
        public int TransactionType { get; set; }
        public double Rate { get; set; }
        public string UserName { get; set; }
        public string RecieptID { get; set; }
        public int IssuanceType { get; set; }
        public int POIssuance { get; set; }
        public int Sheets { get; set; }

        // Presentation property
        public double Balance { get; set; }
    }
}
