using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public interface IManualAttendanceService
    {
        Task<List<EmployeeListItemModel>> GetActiveEmployeesAsync();
        Task<byte[]?> GetEmployeePictureAsync(string empId);
        Task<EmpShiftDefaultTimings> GetEmployeeDefaultTimingsAsync(string empId, System.DateTime date);
        Task<List<EmpTimeDetailRow>> GetEmployeeAttendanceDetailsAsync(string empId, System.DateTime date);
        Task<string?> CheckLeaveOrHolidayAsync(string empId, System.DateTime date);
        Task<bool> SaveManualAttendanceAsync(ManualAttendanceInputDto input);
        Task<bool> DeleteAttendanceAsync(string empId, System.DateTime date);
        Task<int> ReCalculateAttendanceAsync(System.DateTime dateFrom, System.DateTime dateTo, string deptId, string empId);
    }
}
