using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public class AttendanceLedgerService : IAttendanceLedgerService
    {
        private readonly IAttendanceLedgerDataAccess _dataAccess;
        private readonly IDBHelperService _dbHelperService;

        public AttendanceLedgerService(IAttendanceLedgerDataAccess dataAccess, IDBHelperService dbHelperService)
        {
            _dataAccess = dataAccess;
            _dbHelperService = dbHelperService;
        }

        public Task<List<DepartmentListItemModel>> GetDepartmentsAsync() => _dataAccess.GetDepartmentsAsync();

        public Task<List<EmployeeListItemModel>> GetEmployeesAsync(string deptId, DateTime firstDT, DateTime lastDT)
            => _dataAccess.GetEmployeesAsync(deptId, firstDT, lastDT);

        public Task<List<HolidayItemModel>> GetHolidaysAsync(int month, int year) => _dataAccess.GetHolidaysAsync(month, year);

        public Task<List<AttendanceLedgerEntryModel>> GetLedgerEntriesAsync(AttendanceLedgerFilterModel filter)
            => _dataAccess.GetLedgerEntriesAsync(filter);

        public Task<List<AttendanceTimeModel>> GetTimeEntriesAsync(AttendanceLedgerFilterModel filter)
            => _dataAccess.GetTimeEntriesAsync(filter);

        public Task<Dictionary<DateTime, string>> GetShortLeavesAsync(AttendanceLedgerFilterModel filter)
            => _dataAccess.GetShortLeavesAsync(filter);

        public Task PrepareEmpLedgerPrintAsync(List<AttendanceLedgerDayRow> rows, string empId)
            => _dataAccess.PrepareEmpLedgerPrintAsync(rows, empId);

        public Task<string> GetCompanyNameAsync() => _dbHelperService.getGeneralDataValueAsync("Company");
    }
}
