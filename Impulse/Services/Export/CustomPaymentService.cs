using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;
using DataAccessLibrary.Models.ViewModels;

namespace Impulse.Services.Export
{
    public class CustomPaymentService : ICustomPaymentService
    {
        private readonly ICustomPaymentDataAccess _dataAccess;

        public CustomPaymentService(ICustomPaymentDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<List<CustomPaymentStatusModel>> GetCustomPaymentStatusesAsync(string? custCode, int statusIndex)
        {
            return await _dataAccess.GetCustomPaymentStatusesAsync(custCode, statusIndex);
        }

        public async Task<List<AllCustomersBalanceModel>> GetAllCustomersBalanceAsync()
        {
            return await _dataAccess.GetAllCustomersBalanceAsync();
        }

        public async Task<decimal> GetCurrencyExchangeRateAsync(string currency)
        {
            return await _dataAccess.GetCurrencyExchangeRateAsync(currency);
        }

        public async Task<List<GenericDropDownModel>> GetBanksAsync()
        {
            return await _dataAccess.GetBanksAsync();
        }

        public async Task<List<string>> GetCustomInvoicesForPaymentAsync(string custCode)
        {
            return await _dataAccess.GetCustomInvoicesForPaymentAsync(custCode);
        }

        public async Task<decimal> GetCustomInvoiceBalanceAsync(string customInvoice)
        {
            return await _dataAccess.GetCustomInvoiceBalanceAsync(customInvoice);
        }

        public async Task<bool> SaveReceivePaymentAsync(ReceivePaymentModel payment)
        {
            return await _dataAccess.SaveReceivePaymentAsync(payment);
        }

        public async Task<PostPrcModel> GetPrcDetailsAsync(int entryId)
        {
            return await _dataAccess.GetPrcDetailsAsync(entryId);
        }

        public async Task<string> PostPrcAsync(PostPrcModel model)
        {
            return await _dataAccess.PostPrcAsync(model);
        }
    }
}
