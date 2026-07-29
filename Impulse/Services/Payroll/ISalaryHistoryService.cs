using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public interface ISalaryHistoryService
    {
        Task<List<DepartmentListItemModel>> GetDepartmentsAsync();
        Task<List<EmployeeListItemModel>> GetEmployeesAsync(string deptId);
        Task<List<SalaryHistoryRowModel>> GetSalaryHistoryAsync(SalaryHistoryFilterModel filter);
        Task<bool> SaveSalaryChangeAsync(ChangeSalaryRequestModel request);
        Task PrepareAbsentListReportProcAsync(System.DateTime fromDate, System.DateTime toDate);
        Task<string> GetCompanyNameAsync();
    }
}
