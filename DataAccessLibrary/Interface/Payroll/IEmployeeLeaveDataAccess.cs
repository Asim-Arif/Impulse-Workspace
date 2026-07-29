using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface IEmployeeLeaveDataAccess
    {
        /// <summary>
        /// Runs GetLeaveBalances stored procedure and calculates Available, Used, This Month, and Net Balance matrix.
        /// </summary>
        Task<LeaveBalanceMatrixDto> GetLeaveBalancesAsync(string empId, DateTime date);

        /// <summary>
        /// Gets active users list for Leave Authority selection.
        /// </summary>
        Task<List<UserLookupModel>> GetUsersAsync();

        /// <summary>
        /// Checks if employee is currently in probation period on given date.
        /// </summary>
        Task<bool> IsProbationAsync(string empId, DateTime date);

        /// <summary>
        /// Checks if employee attendance or leaves already exist for the selected date range.
        /// </summary>
        Task<bool> HasExistingAttendanceAsync(string empId, DateTime fromDate, DateTime toDate);

        /// <summary>
        /// Saves employee leave or absence entry, deleting existing attendance/leaves if necessary.
        /// </summary>
        Task<bool> SaveEmployeeLeaveAsync(EmployeeLeaveInputDto input);

        /// <summary>
        /// Runs SP_EmpLeaveBalances stored procedure for the monthly leave balance report.
        /// </summary>
        Task ExecutePrintEmpLeaveBalancesAsync(DateTime date);
    }
}
