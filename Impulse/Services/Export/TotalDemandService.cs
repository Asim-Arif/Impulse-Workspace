using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public class TotalDemandService : ITotalDemandService
    {
        private readonly ITotalDemandDataAccess _dataAccess;

        public TotalDemandService(ITotalDemandDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<List<CustomerLookupModel>> GetActiveCustomersAsync()
        {
            return await _dataAccess.GetActiveCustomersAsync();
        }

        public async Task<List<TotalDemandModel>> GetTotalDemandDataAsync(string custCode, DateTime dtFrom, DateTime dtTo)
        {
            return await _dataAccess.GetTotalDemandDataAsync(custCode, dtFrom, dtTo);
        }
    }
}
