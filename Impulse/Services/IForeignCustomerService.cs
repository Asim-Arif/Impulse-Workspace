using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services
{
    public interface IForeignCustomerService
    {
        // Lookups
        Task<List<GenericDropDownModel>> GetCountriesAsync();
        Task<List<GenericDropDownModel>> GetCurrenciesAsync();
        Task<List<GenericDropDownModel>> GetPaymentTermsAsync();
        Task<List<GenericDropDownModel>> GetTradeTermsAsync();
        Task<List<GenericDropDownModel>> GetFinishingQualitiesAsync();
        Task<List<GenericDropDownModel>> GetShippingMethodsAsync();
        Task<List<GenericDropDownModel>> GetCustomerSourcesAsync();
        Task<List<GenericDropDownModel>> GetEventRepetitionsAsync();

        // Customer CRUD
        Task<ForeignCustomerViewModel?> GetCustomerAsync(string custId);
        Task<string> SaveCustomerAsync(ForeignCustomerViewModel model, string custId, string userName, string computerName,
                                       List<int> removedBankIds, List<int> removedAddressIds,
                                       List<int> removedEventIds, List<int> removedInvoiceToIds);
    }
}
