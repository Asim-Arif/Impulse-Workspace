using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;

namespace DataAccessLibrary.Interface.Export
{
    public interface ITotalDemandDataAccess
    {
        Task<List<CustomerLookupModel>> GetActiveCustomersAsync();
        Task<List<TotalDemandModel>> GetTotalDemandDataAsync(string custCode, DateTime dtFrom, DateTime dtTo);
    }
}
