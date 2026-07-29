using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface IBiometricLedgerDataAccess
    {
        /// <summary>
        /// Executes SP_DeptSettings for the given date/dept and returns dept summary rows.
        /// </summary>
        Task<List<BiometricDeptSummaryRow>> GetDeptSummaryAsync(DateTime date, string? deptId);

        /// <summary>
        /// Returns employee attendance rows from VEmpTimes1 RIGHT OUTER JOIN VEmp for the selected date/dept/employee.
        /// </summary>
        Task<List<BiometricEmployeeRow>> GetEmployeeAttendanceAsync(DateTime date, string? deptId, string? empId);

        /// <summary>
        /// Returns all IFace machines (MachineType = 3) from the AttendanceMachine table.
        /// </summary>
        Task<List<BiometricMachineItem>> GetIFaceMachinesAsync();

        /// <summary>
        /// Calls SP_PrintDailyAbsentLeaveDetail to prepare the staging table for the daily attendance report.
        /// </summary>
        Task ExecutePrintDailyAbsentLeaveDetailAsync(DateTime date);

        /// <summary>
        /// Calls SP_PrintDailyAbsentLeaveSummary to prepare the staging table for the absent/leave summary report.
        /// </summary>
        Task ExecutePrintDailyAbsentLeaveSummaryAsync(DateTime date);
    }
}
