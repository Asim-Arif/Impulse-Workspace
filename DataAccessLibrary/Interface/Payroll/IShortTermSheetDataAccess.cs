using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface IShortTermSheetDataAccess
    {
        Task<List<ShortTermSheetItemViewModel>> GetShortTermSheetDataAsync(DateTime date, string deptId);
        Task SaveShortTermSheetEntriesAsync(DateTime date, IEnumerable<ShortTermSheetSaveDto> items);
        Task<ShortTermPostingStatusModel> GetPostingStatusAsync(DateTime date, string deptId);
        Task SavePostingRecordAsync(DateTime date, string deptId, string vchrNo, string bankVchrNo, string chqNo, DateTime chqDT, int postingType, string userName, string machineName, IEnumerable<ShortTermSheetItemViewModel> items);
    }
}

