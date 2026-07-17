using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public class TotalExportService : ITotalExportService
    {
        private readonly ITotalExportDataAccess _dataAccess;

        public TotalExportService(ITotalExportDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<List<CustomerLookupModel>> GetActiveCustomersAsync()
        {
            return await _dataAccess.GetActiveCustomersAsync();
        }

        public async Task<List<TotalExportModel>> GetTotalExportDataAsync(string custCode, DateTime dtFrom, DateTime dtTo)
        {
            return await _dataAccess.GetTotalExportDataAsync(custCode, dtFrom, dtTo);
        }
    }
}
