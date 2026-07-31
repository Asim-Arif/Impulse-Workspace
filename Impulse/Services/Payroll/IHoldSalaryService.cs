using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public interface IHoldSalaryService
    {
        Task<HoldSalarySummaryModel> GetEmployeeHoldSalaryInfoAsync(string empId, int year, int month);
        Task<List<HoldSalaryListItemModel>> GetHoldSalariesAsync(string? deptId, int year, int month);
        Task<bool> SaveHoldSalaryAsync(HoldSalaryDto dto);
        Task<bool> DeleteHoldSalaryAsync(int entryId);
    }
}
