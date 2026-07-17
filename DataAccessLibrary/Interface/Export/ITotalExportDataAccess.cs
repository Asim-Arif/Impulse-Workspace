using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;

namespace DataAccessLibrary.Interface.Export
{
    public interface ITotalExportDataAccess
    {
        Task<List<CustomerLookupModel>> GetActiveCustomersAsync();
        Task<List<TotalExportModel>> GetTotalExportDataAsync(string custCode, DateTime dtFrom, DateTime dtTo);
    }
}
