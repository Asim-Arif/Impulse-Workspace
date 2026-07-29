using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public class BiometricLedgerService : IBiometricLedgerService
    {
        private readonly IBiometricLedgerDataAccess _dac;

        public BiometricLedgerService(IBiometricLedgerDataAccess dac)
        {
            _dac = dac;
        }

        public Task<List<BiometricDeptSummaryRow>> GetDeptSummaryAsync(DateTime date, string? deptId)
            => _dac.GetDeptSummaryAsync(date, deptId);

        public Task<List<BiometricEmployeeRow>> GetEmployeeAttendanceAsync(DateTime date, string? deptId, string? empId)
            => _dac.GetEmployeeAttendanceAsync(date, deptId, empId);

        public Task<List<BiometricMachineItem>> GetIFaceMachinesAsync()
            => _dac.GetIFaceMachinesAsync();

        public Task ExecutePrintDailyAbsentLeaveDetailAsync(DateTime date)
            => _dac.ExecutePrintDailyAbsentLeaveDetailAsync(date);

        public Task ExecutePrintDailyAbsentLeaveSummaryAsync(DateTime date)
            => _dac.ExecutePrintDailyAbsentLeaveSummaryAsync(date);
    }
}
