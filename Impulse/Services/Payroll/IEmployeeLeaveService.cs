using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public interface IEmployeeLeaveService
    {
        Task<LeaveBalanceMatrixDto> GetLeaveBalancesAsync(string empId, DateTime date);
        Task<List<UserLookupModel>> GetUsersAsync();
        Task<bool> IsProbationAsync(string empId, DateTime date);
        Task<bool> HasExistingAttendanceAsync(string empId, DateTime fromDate, DateTime toDate);
        Task<bool> SaveEmployeeLeaveAsync(EmployeeLeaveInputDto input);
        Task ExecutePrintEmpLeaveBalancesAsync(DateTime date);
    }
}
