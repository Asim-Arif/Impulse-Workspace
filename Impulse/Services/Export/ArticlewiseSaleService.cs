using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public class ArticlewiseSaleService : IArticlewiseSaleService
    {
        private readonly IArticlewiseSaleDataAccess _dataAccess;

        public ArticlewiseSaleService(IArticlewiseSaleDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<List<ItemLookupModel>> GetArticlesAsync()
        {
            return await _dataAccess.GetArticlesAsync();
        }

        public async Task<List<CustomerLookupModel>> GetActiveCustomersAsync()
        {
            return await _dataAccess.GetActiveCustomersAsync();
        }

        public async Task<List<string>> GetCountriesAsync(string custCode)
        {
            return await _dataAccess.GetCountriesAsync(custCode);
        }

        public async Task<List<ArticlewiseSaleModel>> GetArticlewiseSalesDataAsync(string itemId, string custCode, string country, bool useDateRange, DateTime dtFrom, DateTime dtTo)
        {
            return await _dataAccess.GetArticlewiseSalesDataAsync(itemId, custCode, country, useDateRange, dtFrom, dtTo);
        }
    }
}
