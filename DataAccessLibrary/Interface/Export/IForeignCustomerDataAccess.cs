using DataAccessLibrary.Models.ViewModels.Export;

namespace DataAccessLibrary.Interface.Export
{
    public interface IForeignCustomerDataAccess
    {
        // custId = CustCode + Country composite key (e.g. "C001Germany")
        Task<ForeignCustomerViewModel?> GetCustomerAsync(string custId);
        Task<string> InsertCustomerAsync(ForeignCustomerViewModel model, string userName, string computerName);
        Task UpdateCustomerAsync(ForeignCustomerViewModel model, string custId);
        Task SaveBanksAsync(ForeignCustomerViewModel model, string custCode, string country, List<int> removedBankIds);
        Task SaveShippingAddressesAsync(ForeignCustomerViewModel model, string custCode, string country, List<int> removedIds);
        Task SaveEventsAsync(ForeignCustomerViewModel model, string custCode, string country, List<int> removedIds);
        Task SaveInvoiceToAsync(ForeignCustomerViewModel model, string custCode, string country, List<int> removedIds);
    }
}
