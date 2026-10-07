using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public class ManualAttendanceService : IManualAttendanceService
    {
        private readonly IManualAttendanceDataAccess _dataAccess;

        public ManualAttendanceService(IManualAttendanceDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<EmployeeListItemModel>> GetActiveEmployeesAsync() => _dataAccess.GetActiveEmployeesAsync();

        public Task<byte[]?> GetEmployeePictureAsync(string empId) => _dataAccess.GetEmployeePictureAsync(empId);

        public Task<EmpShiftDefaultTimings> GetEmployeeDefaultTimingsAsync(string empId, System.DateTime date) => _dataAccess.GetEmployeeDefaultTimingsAsync(empId, date);

        public Task<List<EmpTimeDetailRow>> GetEmployeeAttendanceDetailsAsync(string empId, System.DateTime date) => _dataAccess.GetEmployeeAttendanceDetailsAsync(empId, date);

        public Task<string?> CheckLeaveOrHolidayAsync(string empId, System.DateTime date) => _dataAccess.CheckLeaveOrHolidayAsync(empId, date);

        public async Task<bool> SaveManualAttendanceAsync(ManualAttendanceInputDto input)
        {
            if (await _dataAccess.IsSalaryFinalizedAsync(input.DT.Year, input.DT.Month))
            {
                throw new System.InvalidOperationException($"Salary for {input.DT.Year}-{input.DT.Month:D2} has already been finalized. Attendance cannot be modified.");
            }
            return await _dataAccess.SaveManualAttendanceAsync(input);
        }

        public async Task<bool> DeleteAttendanceAsync(string empId, System.DateTime date)
        {
            if (await _dataAccess.IsSalaryFinalizedAsync(date.Year, date.Month))
            {
                throw new System.InvalidOperationException($"Salary for {date.Year}-{date.Month:D2} has already been finalized. Attendance cannot be modified.");
            }
            return await _dataAccess.DeleteAttendanceAsync(empId, date);
        }

        public Task<int> ReCalculateAttendanceAsync(System.DateTime dateFrom, System.DateTime dateTo, string deptId, string empId) => _dataAccess.ReCalculateAttendanceAsync(dateFrom, dateTo, deptId, empId);

        public Task<bool> IsSalaryFinalizedAsync(int year, int month) => _dataAccess.IsSalaryFinalizedAsync(year, month);
    }
}
