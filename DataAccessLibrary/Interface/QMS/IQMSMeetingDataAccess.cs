using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.QMS;

namespace DataAccessLibrary.Interface.QMS
{
    public interface IQMSMeetingDataAccess
    {
        Task<List<QMSMeetingListItemModel>> GetMeetingsAsync(QMSMeetingFilterDto filter);
        Task<QMSMeetingModel?> GetMeetingByIdAsync(int entryId);
        Task<int> SaveMeetingAsync(QMSMeetingModel model, string currentUserName, string? machineName);
        Task<bool> DeleteMeetingAsync(int entryId);
        Task<string> GetNextMeetingCodeAsync();
        Task<List<string>> GetMeetingTypesAsync();
        Task<List<EmployeeDropdownDto>> GetEmployeesAsync();
    }
}

