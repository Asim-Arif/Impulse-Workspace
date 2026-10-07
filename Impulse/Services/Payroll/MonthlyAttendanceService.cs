using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public class MonthlyAttendanceService : IMonthlyAttendanceService
    {
        private readonly IMonthlyAttendanceDataAccess _dac;

        public MonthlyAttendanceService(IMonthlyAttendanceDataAccess dac)
        {
            _dac = dac;
        }

        public Task<List<MonthlyAttendanceDayRow>> GetMonthlyAttendanceAsync(string empId, int year, int month)
            => _dac.GetMonthlyAttendanceAsync(empId, year, month);

        public async Task<bool> SaveMonthlyAttendanceAsync(MonthlyAttendanceSaveDto input)
        {
            if (await _dac.IsSalaryFinalizedAsync(input.Year, input.Month))
            {
                throw new InvalidOperationException($"Salary for {input.Year}-{input.Month:D2} has already been finalized. Attendance cannot be modified.");
            }
            return await _dac.SaveMonthlyAttendanceAsync(input);
        }

        public async Task<bool> ClearDateAttendanceAsync(string empId, DateTime date)
        {
            if (await _dac.IsSalaryFinalizedAsync(date.Year, date.Month))
            {
                throw new InvalidOperationException($"Salary for {date.Year}-{date.Month:D2} has already been finalized. Attendance cannot be modified.");
            }
            return await _dac.ClearDateAttendanceAsync(empId, date);
        }

        public Task<bool> IsSalaryFinalizedAsync(int year, int month)
            => _dac.IsSalaryFinalizedAsync(year, month);
    }
}
