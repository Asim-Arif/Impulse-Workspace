using DataAccessLibrary.Interface.Payroll;
using System;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public class TakeAttendanceService : ITakeAttendanceService
    {
        private readonly ITakeAttendanceDataAccess _dataAccess;

        public TakeAttendanceService(ITakeAttendanceDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<bool> ProcessPunchAsync(string empId, DateTime date, string punchTimeStr, string userName = "System", string computerName = "Unknown")
        {
            return _dataAccess.ProcessPunchAsync(empId, date, punchTimeStr, userName, computerName);
        }

        public Task<int> ReCalculateAttendanceAsync(DateTime dateFrom, DateTime dateTo, string deptId, string empId)
        {
            return _dataAccess.ReCalculateAttendanceAsync(dateFrom, dateTo, deptId, empId);
        }

        public Task<int> AutoGenerateAttendanceAsync(string deptId, string empId, DateTime dateFrom, DateTime dateTo)
        {
            return _dataAccess.AutoGenerateAttendanceAsync(deptId, empId, dateFrom, dateTo);
        }

        public Task<int> CompleteMissingAttendanceAsync(string deptId, string empId, DateTime dateFrom, DateTime dateTo)
        {
            return _dataAccess.CompleteMissingAttendanceAsync(deptId, empId, dateFrom, dateTo);
        }
    }
}
