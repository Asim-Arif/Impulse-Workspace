using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public interface IExportPerformanceReportService
    {
        Task<List<CustomerLookupModel>> GetCustomersAsync();
        Task<List<ExportItemGroupLookupModel>> GetItemGroupsAsync();
        Task<List<ExportItemCategoryLookupModel>> GetItemCategoriesAsync();
        Task<List<ExportMainGroupLookupModel>> GetMainGroupsAsync();
    }
}
