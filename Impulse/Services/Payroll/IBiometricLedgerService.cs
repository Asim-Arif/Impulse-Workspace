using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public interface IBiometricLedgerService
    {
        Task<List<BiometricDeptSummaryRow>> GetDeptSummaryAsync(DateTime date, string? deptId);
        Task<List<BiometricEmployeeRow>> GetEmployeeAttendanceAsync(DateTime date, string? deptId, string? empId);
        Task<List<BiometricMachineItem>> GetIFaceMachinesAsync();
        Task ExecutePrintDailyAbsentLeaveDetailAsync(DateTime date);
        Task ExecutePrintDailyAbsentLeaveSummaryAsync(DateTime date);
    }
}
