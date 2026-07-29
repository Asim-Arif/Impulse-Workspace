using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface IMonthlyAttendanceDataAccess
    {
        /// <summary>
        /// Gets full month attendance day rows for a specific employee, including punches, status, holidays, & OT.
        /// </summary>
        Task<List<MonthlyAttendanceDayRow>> GetMonthlyAttendanceAsync(string empId, int year, int month);

        /// <summary>
        /// Saves month attendance rows by processing punches through ITakeAttendanceDataAccess and applying breaks.
        /// </summary>
        Task<bool> SaveMonthlyAttendanceAsync(MonthlyAttendanceSaveDto input);

        /// <summary>
        /// Clears attendance & punches for a specific day.
        /// </summary>
        Task<bool> ClearDateAttendanceAsync(string empId, DateTime date);
    }
}
