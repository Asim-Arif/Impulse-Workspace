using System;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public interface ITakeAttendanceService
    {
        Task<bool> ProcessPunchAsync(string empId, DateTime date, string punchTimeStr, string userName = "System", string computerName = "Unknown");
        Task<int> ReCalculateAttendanceAsync(DateTime dateFrom, DateTime dateTo, string deptId, string empId);
        Task<int> AutoGenerateAttendanceAsync(string deptId, string empId, DateTime dateFrom, DateTime dateTo);
        Task<int> CompleteMissingAttendanceAsync(string deptId, string empId, DateTime dateFrom, DateTime dateTo);
    }
}
