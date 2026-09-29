using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;

namespace DataAccessLibrary.Interface.Production
{
    public interface IEmployeePerformanceDataAccess
    {
        Task<List<DepartmentLookupItem>> GetDepartmentsLookupAsync();
        Task<List<ProcessLookupItem>> GetProcessesLookupAsync();
        Task<List<PerformanceEmployeeLookupItem>> GetEmployeesLookupAsync(string? deptId = null);
        Task<EmployeePerformanceDashboardDto> GetPerformanceDashboardAsync(EmployeePerformanceFilter filter);
        Task<List<EmployeeLotDetailDto>> GetEmployeeLotsDetailAsync(string empId, DateTime dtFrom, DateTime dtTo);
    }
}
