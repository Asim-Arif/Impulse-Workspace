using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface IManualAttendanceDataAccess
    {
        Task<List<EmployeeListItemModel>> GetActiveEmployeesAsync();
        Task<byte[]?> GetEmployeePictureAsync(string empId);
        Task<EmpShiftDefaultTimings> GetEmployeeDefaultTimingsAsync(string empId, DateTime date);
        Task<List<EmpTimeDetailRow>> GetEmployeeAttendanceDetailsAsync(string empId, DateTime date);
        Task<string?> CheckLeaveOrHolidayAsync(string empId, DateTime date);
        Task<bool> SaveManualAttendanceAsync(ManualAttendanceInputDto input);
        Task<bool> DeleteAttendanceAsync(string empId, DateTime date);
        Task<int> ReCalculateAttendanceAsync(DateTime dateFrom, DateTime dateTo, string deptId, string empId);
    }
}
