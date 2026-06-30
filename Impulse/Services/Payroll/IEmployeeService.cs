using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public interface IEmployeeService
    {
        Task<List<DepartmentListItemModel>> GetDepartmentsAsync(bool showInactive, string? selectedDeptId = null);
        Task<List<EmployeeListItemModel>> GetEmployeesAsync(string deptId, bool showInactive, bool tempEmpOnly, string? selectedEmpId = null);
        Task<byte[]?> GetEmployeePictureAsync(string empId);
        Task SetEmployeeActiveStateAsync(string empId, bool active, DateTime? date = null, string? remarks = null);
        Task<string> GetCompanyNameAsync();
        Task SaveEmployeeAsync(EmployeeDto employee, bool isAdd);
        Task<string> GetNextEmpIDAsync(string deptId);
        Task<List<ProcessLookupModel>> GetProcessesAsync();
        Task<EmployeeDto?> GetEmployeeByIdAsync(string empId);
        Task<List<GroupLookupModel>> GetGroupsAsync();
    }
}
