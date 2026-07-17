using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public interface ITotalDemandService
    {
        Task<List<CustomerLookupModel>> GetActiveCustomersAsync();
        Task<List<TotalDemandModel>> GetTotalDemandDataAsync(string custCode, DateTime dtFrom, DateTime dtTo);
    }
}
