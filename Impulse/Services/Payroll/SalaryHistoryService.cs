using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public class SalaryHistoryService : ISalaryHistoryService
    {
        private readonly ISalaryHistoryDataAccess _dataAccess;
        private readonly IDBHelperService _dbHelperService;

        public SalaryHistoryService(ISalaryHistoryDataAccess dataAccess, IDBHelperService dbHelperService)
        {
            _dataAccess = dataAccess;
            _dbHelperService = dbHelperService;
        }

        public Task<List<DepartmentListItemModel>> GetDepartmentsAsync() => _dataAccess.GetDepartmentsAsync();

        public Task<List<EmployeeListItemModel>> GetEmployeesAsync(string deptId) => _dataAccess.GetEmployeesAsync(deptId);

        public Task<List<SalaryHistoryRowModel>> GetSalaryHistoryAsync(SalaryHistoryFilterModel filter) => _dataAccess.GetSalaryHistoryAsync(filter);

        public Task<bool> SaveSalaryChangeAsync(ChangeSalaryRequestModel request) => _dataAccess.SaveSalaryChangeAsync(request);

        public Task PrepareAbsentListReportProcAsync(System.DateTime fromDate, System.DateTime toDate) => _dataAccess.PrepareAbsentListReportProcAsync(fromDate, toDate);

        public Task<string> GetCompanyNameAsync() => _dbHelperService.getGeneralDataValueAsync("Company");
    }
}
