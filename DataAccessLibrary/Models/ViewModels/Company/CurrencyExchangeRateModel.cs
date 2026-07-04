using System;

namespace DataAccessLibrary.Models.ViewModels.Company
{
    public class CurrencyExchangeRateModel
    {
        public int EntryID { get; set; }
        public string Currency { get; set; } = string.Empty;
        public float ExchRate { get; set; }
        
        // UI helper properties
        public bool IsEditing { get; set; }
        public float EditRate { get; set; }
    }
}
