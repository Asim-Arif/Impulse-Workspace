namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class AllCustomersBalanceModel
    {
        public string Currency { get; set; } = string.Empty;
        public decimal Balance { get; set; }
        public decimal ExchangeRate { get; set; }
        public decimal BalancePK => Balance * ExchangeRate;
    }
}
