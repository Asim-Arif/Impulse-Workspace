using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public class ExportPerformanceReportService : IExportPerformanceReportService
    {
        private readonly IExportPerformanceReportDataAccess _dataAccess;

        public ExportPerformanceReportService(IExportPerformanceReportDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<CustomerLookupModel>> GetCustomersAsync() => _dataAccess.GetCustomersAsync();
        public Task<List<ExportItemGroupLookupModel>> GetItemGroupsAsync() => _dataAccess.GetItemGroupsAsync();
        public Task<List<ExportItemCategoryLookupModel>> GetItemCategoriesAsync() => _dataAccess.GetItemCategoriesAsync();
        public Task<List<ExportMainGroupLookupModel>> GetMainGroupsAsync() => _dataAccess.GetMainGroupsAsync();
    }
}
