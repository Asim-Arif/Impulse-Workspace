using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public interface ITotalExportService
    {
        Task<List<CustomerLookupModel>> GetActiveCustomersAsync();
        Task<List<TotalExportModel>> GetTotalExportDataAsync(string custCode, DateTime dtFrom, DateTime dtTo);
    }
}
