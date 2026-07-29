using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public interface IMonthlyAttendanceService
    {
        Task<List<MonthlyAttendanceDayRow>> GetMonthlyAttendanceAsync(string empId, int year, int month);
        Task<bool> SaveMonthlyAttendanceAsync(MonthlyAttendanceSaveDto input);
        Task<bool> ClearDateAttendanceAsync(string empId, DateTime date);
    }
}
