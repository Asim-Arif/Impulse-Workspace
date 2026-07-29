using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public class EmployeeLeaveService : IEmployeeLeaveService
    {
        private readonly IEmployeeLeaveDataAccess _dac;

        public EmployeeLeaveService(IEmployeeLeaveDataAccess dac)
        {
            _dac = dac;
        }

        public Task<LeaveBalanceMatrixDto> GetLeaveBalancesAsync(string empId, DateTime date)
            => _dac.GetLeaveBalancesAsync(empId, date);

        public Task<List<UserLookupModel>> GetUsersAsync()
            => _dac.GetUsersAsync();

        public Task<bool> IsProbationAsync(string empId, DateTime date)
            => _dac.IsProbationAsync(empId, date);

        public Task<bool> HasExistingAttendanceAsync(string empId, DateTime fromDate, DateTime toDate)
            => _dac.HasExistingAttendanceAsync(empId, fromDate, toDate);

        public Task<bool> SaveEmployeeLeaveAsync(EmployeeLeaveInputDto input)
            => _dac.SaveEmployeeLeaveAsync(input);

        public Task ExecutePrintEmpLeaveBalancesAsync(DateTime date)
            => _dac.ExecutePrintEmpLeaveBalancesAsync(date);
    }
}
