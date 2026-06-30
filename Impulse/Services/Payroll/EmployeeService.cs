using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public class EmployeeService : IEmployeeService
    {
        private readonly IEmployeeDataAccess _dataAccess;
        private readonly IAuditService _auditService;
        private readonly IDBHelperService _dbHelperService;

        public EmployeeService(IEmployeeDataAccess dataAccess, IAuditService auditService, IDBHelperService dbHelperService)
        {
            _dataAccess = dataAccess;
            _auditService = auditService;
            _dbHelperService = dbHelperService;
        }

        public async Task<List<DepartmentListItemModel>> GetDepartmentsAsync(bool showInactive, string? selectedDeptId = null)
        {
            return await _dataAccess.GetDepartmentsAsync(showInactive, selectedDeptId);
        }

        public async Task<List<EmployeeListItemModel>> GetEmployeesAsync(string deptId, bool showInactive, bool tempEmpOnly, string? selectedEmpId = null)
        {
            return await _dataAccess.GetEmployeesAsync(deptId, showInactive, tempEmpOnly, selectedEmpId);
        }

        public async Task<byte[]?> GetEmployeePictureAsync(string empId)
        {
            return await _dataAccess.GetEmployeePictureAsync(empId);
        }

        public async Task SetEmployeeActiveStateAsync(string empId, bool active, DateTime? date = null, string? remarks = null)
        {
            string userName = _auditService.GetCurrentUserName();
            string machineName = _auditService.GetClientIpAddress();
            await _dataAccess.SetEmployeeActiveStateAsync(empId, active, date, remarks, userName, machineName);
        }

        public async Task<string> GetCompanyNameAsync()
        {
            return await _dbHelperService.getGeneralDataValueAsync("Company");
        }

        public async Task SaveEmployeeAsync(EmployeeDto employee, bool isAdd)
        {
            string userName = _auditService.GetCurrentUserName();
            string machineName = _auditService.GetClientIpAddress();
            await _dataAccess.SaveEmployeeAsync(employee, isAdd, userName, machineName);
        }

        public async Task<string> GetNextEmpIDAsync(string deptId)
        {
            return await _dataAccess.GetNextEmpIDAsync(deptId);
        }

        public async Task<List<ProcessLookupModel>> GetProcessesAsync()
        {
            return await _dataAccess.GetProcessesAsync();
        }

        public async Task<EmployeeDto?> GetEmployeeByIdAsync(string empId)
        {
            return await _dataAccess.GetEmployeeByIdAsync(empId);
        }

        public async Task<List<GroupLookupModel>> GetGroupsAsync()
        {
            return await _dataAccess.GetGroupsAsync();
        }
    }
}
