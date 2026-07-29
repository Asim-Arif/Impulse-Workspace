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

        public Task<bool> SaveMonthlyAttendanceAsync(MonthlyAttendanceSaveDto input)
            => _dac.SaveMonthlyAttendanceAsync(input);

        public Task<bool> ClearDateAttendanceAsync(string empId, DateTime date)
            => _dac.ClearDateAttendanceAsync(empId, date);
    }
}
