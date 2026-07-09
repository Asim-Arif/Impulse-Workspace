using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;

namespace DataAccessLibrary.Interface.Export
{
    public interface ICustomerQuotationDataAccess
    {
        Task<List<CustomerQuotationListItemModel>> GetQuotationListAsync(
            DateTime dtFrom,
            DateTime dtTo,
            string custCode,
            string country);

        Task<CustomerQuotationViewModel?> GetQuotationAsync(int quotationNo);
        Task<bool> SaveQuotationAsync(CustomerQuotationViewModel quotation, List<int> deletedItemIds, List<int> deletedChargeIds);
        Task<int> GetNextQuotationNoAsync();
        Task<bool> DeleteQuotationAsync(int quotationNo);
    }
}
