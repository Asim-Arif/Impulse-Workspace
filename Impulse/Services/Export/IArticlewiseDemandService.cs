using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public interface IArticlewiseDemandService
    {
        Task<List<ItemLookupModel>> GetArticlesAsync();
        Task<List<CustomerLookupModel>> GetActiveCustomersAsync();
        Task<List<string>> GetCountriesAsync(string custCode);
        Task<List<ArticlewiseDemandModel>> GetArticlewiseDemandDataAsync(string itemId, string custCode, string country, bool useDateRange, DateTime dtFrom, DateTime dtTo);
    }
}
