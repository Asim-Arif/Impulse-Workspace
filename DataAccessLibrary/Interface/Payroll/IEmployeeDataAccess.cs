using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface IEmployeeDataAccess
    {
        Task<List<DepartmentListItemModel>> GetDepartmentsAsync(bool showInactive, string? selectedDeptId = null);
        Task<List<EmployeeListItemModel>> GetEmployeesAsync(string deptId, bool showInactive, bool tempEmpOnly, string? selectedEmpId = null);
        Task<byte[]?> GetEmployeePictureAsync(string empId);
        Task SetEmployeeActiveStateAsync(string empId, bool active, DateTime? date = null, string? remarks = null, string? userName = null, string? machineName = null);
        Task SaveEmployeeAsync(EmployeeDto employee, bool isAdd, string userName, string machineName);
        Task<string> GetNextEmpIDAsync(string deptId);
    }
}
