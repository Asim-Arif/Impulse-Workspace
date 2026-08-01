using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public interface IPayrollSettingsService
    {
        Task<GeneralSettingsModel> GetSettingsByDeptAsync(string deptId);
        Task<bool> SaveSettingsAsync(GeneralSettingsModel model);
        Task<bool> UpdateAllDepartmentTimingsAsync(string enterTime, string exitTime, string lunchInTime, string lunchOutTime, string fridayBreakStart, string fridayBreakEnd, string fridayCheckOut);
        Task<List<EmpAttCategoryModel>> GetAttendanceCategoriesAsync();
        Task<List<CategoryRuleItem>> GetCategoryRulesByIdAsync(int categoryId);
    }
}
