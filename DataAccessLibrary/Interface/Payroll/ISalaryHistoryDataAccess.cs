using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface ISalaryHistoryDataAccess
    {
        Task<List<DepartmentListItemModel>> GetDepartmentsAsync();
        Task<List<EmployeeListItemModel>> GetEmployeesAsync(string deptId);
        Task<List<SalaryHistoryRowModel>> GetSalaryHistoryAsync(SalaryHistoryFilterModel filter);
        Task<bool> SaveSalaryChangeAsync(ChangeSalaryRequestModel request);
        Task PrepareAbsentListReportProcAsync(DateTime fromDate, DateTime toDate);
    }
}
