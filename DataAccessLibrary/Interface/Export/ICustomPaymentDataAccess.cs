using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Export;

namespace DataAccessLibrary.Interface.Export
{
    public interface ICustomPaymentDataAccess
    {
        Task<List<CustomPaymentStatusModel>> GetCustomPaymentStatusesAsync(string? custCode, int statusIndex);
        Task<List<AllCustomersBalanceModel>> GetAllCustomersBalanceAsync();
        Task<decimal> GetCurrencyExchangeRateAsync(string currency);
        Task<List<GenericDropDownModel>> GetBanksAsync();
        Task<List<GenericDropDownModel>> GetPrcBanksAsync();
        Task<List<string>> GetCustomInvoicesForPaymentAsync(string custCode);
        Task<decimal> GetCustomInvoiceBalanceAsync(string customInvoice);
        Task<bool> SaveReceivePaymentAsync(ReceivePaymentModel payment);
        Task<PostPrcModel> GetPrcDetailsAsync(int entryId);
        Task<string> PostPrcAsync(PostPrcModel model);
    }
}
