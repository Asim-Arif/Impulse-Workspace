using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Export;

namespace DataAccessLibrary.Interface.Export
{
    public interface IArticlewiseSaleDataAccess
    {
        Task<List<ItemLookupModel>> GetArticlesAsync();
        Task<List<CustomerLookupModel>> GetActiveCustomersAsync();
        Task<List<string>> GetCountriesAsync(string custCode);
        Task<List<ArticlewiseSaleModel>> GetArticlewiseSalesDataAsync(string itemId, string custCode, string country, bool useDateRange, DateTime dtFrom, DateTime dtTo);
    }
}
