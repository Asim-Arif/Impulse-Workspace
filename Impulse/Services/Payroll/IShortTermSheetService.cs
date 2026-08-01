using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public interface IShortTermSheetService
    {
        Task<List<ShortTermSheetItemViewModel>> GetShortTermSheetDataAsync(DateTime date, string deptId);
        Task SaveShortTermSheetEntriesAsync(DateTime date, IEnumerable<ShortTermSheetSaveDto> items);
        Task<ShortTermPostingStatusModel> GetPostingStatusAsync(DateTime date, string deptId);
        Task<string> PostVouchersAndAdvancesAsync(ShortTermPostingDto dto);
    }
}

