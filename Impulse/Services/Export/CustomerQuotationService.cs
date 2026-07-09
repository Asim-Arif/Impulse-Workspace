using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public class CustomerQuotationService : ICustomerQuotationService
    {
        private readonly ICustomerQuotationDataAccess _dataAccess;

        public CustomerQuotationService(ICustomerQuotationDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<List<CustomerQuotationListItemModel>> GetQuotationListAsync(DateTime dtFrom, DateTime dtTo, string custCode, string country)
        {
            return await _dataAccess.GetQuotationListAsync(dtFrom, dtTo, custCode, country);
        }

        public async Task<CustomerQuotationViewModel?> GetQuotationAsync(int quotationNo)
        {
            return await _dataAccess.GetQuotationAsync(quotationNo);
        }

        public async Task<bool> SaveQuotationAsync(CustomerQuotationViewModel quotation, List<int> deletedItemIds, List<int> deletedChargeIds)
        {
            return await _dataAccess.SaveQuotationAsync(quotation, deletedItemIds, deletedChargeIds);
        }

        public async Task<int> GetNextQuotationNoAsync()
        {
            return await _dataAccess.GetNextQuotationNoAsync();
        }

        public async Task<bool> DeleteQuotationAsync(int quotationNo)
        {
            return await _dataAccess.DeleteQuotationAsync(quotationNo);
        }
    }
}
