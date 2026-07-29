using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface IAttendanceLedgerDataAccess
    {
        Task<List<DepartmentListItemModel>> GetDepartmentsAsync();
        Task<List<EmployeeListItemModel>> GetEmployeesAsync(string deptId, DateTime firstDT, DateTime lastDT);
        Task<List<HolidayItemModel>> GetHolidaysAsync(int month, int year);
        Task<List<AttendanceLedgerEntryModel>> GetLedgerEntriesAsync(AttendanceLedgerFilterModel filter);
        Task<List<AttendanceTimeModel>> GetTimeEntriesAsync(AttendanceLedgerFilterModel filter);
        Task<Dictionary<DateTime, string>> GetShortLeavesAsync(AttendanceLedgerFilterModel filter);
        Task PrepareEmpLedgerPrintAsync(List<AttendanceLedgerDayRow> rows, string empId);
    }
}
