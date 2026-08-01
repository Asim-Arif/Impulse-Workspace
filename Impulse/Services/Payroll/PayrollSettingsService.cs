using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.DAC.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public class PayrollSettingsService : IPayrollSettingsService
    {
        private readonly IPayrollSettingsDataAccess _dac;

        public PayrollSettingsService(IPayrollSettingsDataAccess dac)
        {
            _dac = dac;
        }

        public async Task<GeneralSettingsModel> GetSettingsByDeptAsync(string deptId)
        {
            return await _dac.GetSettingsByDeptAsync(deptId);
        }

        public async Task<bool> SaveSettingsAsync(GeneralSettingsModel model)
        {
            return await _dac.SaveSettingsAsync(model);
        }

        public async Task<bool> UpdateAllDepartmentTimingsAsync(string enterTime, string exitTime, string lunchInTime, string lunchOutTime, string fridayBreakStart, string fridayBreakEnd, string fridayCheckOut)
        {
            return await _dac.UpdateAllDepartmentTimingsAsync(enterTime, exitTime, lunchInTime, lunchOutTime, fridayBreakStart, fridayBreakEnd, fridayCheckOut);
        }

        public async Task<List<EmpAttCategoryModel>> GetAttendanceCategoriesAsync()
        {
            return await _dac.GetAttendanceCategoriesAsync();
        }

        public async Task<List<CategoryRuleItem>> GetCategoryRulesByIdAsync(int categoryId)
        {
            return await _dac.GetCategoryRulesByIdAsync(categoryId);
        }
    }
}
