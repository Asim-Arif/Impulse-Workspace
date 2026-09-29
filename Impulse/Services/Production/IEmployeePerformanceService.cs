using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;

namespace Impulse.Services.Production
{
    public interface IEmployeePerformanceService
    {
        Task<List<DepartmentLookupItem>> GetDepartmentsLookupAsync();
        Task<List<ProcessLookupItem>> GetProcessesLookupAsync();
        Task<List<PerformanceEmployeeLookupItem>> GetEmployeesLookupAsync(int? deptId = null);
        Task<EmployeePerformanceDashboardDto> GetPerformanceDashboardAsync(EmployeePerformanceFilter filter);
        Task<List<EmployeeLotDetailDto>> GetEmployeeLotsDetailAsync(string empId, DateTime dtFrom, DateTime dtTo);
    }
}
