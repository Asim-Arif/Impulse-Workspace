using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;

namespace DataAccessLibrary.Interface.Export
{
    public interface IExportPerformanceReportDataAccess
    {
        Task<List<CustomerLookupModel>> GetCustomersAsync();
        Task<List<ExportItemGroupLookupModel>> GetItemGroupsAsync();
        Task<List<ExportItemCategoryLookupModel>> GetItemCategoriesAsync();
        Task<List<ExportMainGroupLookupModel>> GetMainGroupsAsync();
    }
}
