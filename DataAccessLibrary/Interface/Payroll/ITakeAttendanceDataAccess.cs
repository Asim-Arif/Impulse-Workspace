using System;
using System.Data;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface ITakeAttendanceDataAccess
    {
        Task<bool> ProcessPunchAsync(string empId, DateTime date, string punchTimeStr, string userName = "System", string computerName = "Unknown", IDbConnection? db = null, IDbTransaction? trans = null);
        Task<int> ReCalculateAttendanceAsync(DateTime dateFrom, DateTime dateTo, string deptId, string empId);
        Task<int> AutoGenerateAttendanceAsync(string deptId, string empId, DateTime dateFrom, DateTime dateTo);
        Task<int> CompleteMissingAttendanceAsync(string deptId, string empId, DateTime dateFrom, DateTime dateTo);
    }
}
