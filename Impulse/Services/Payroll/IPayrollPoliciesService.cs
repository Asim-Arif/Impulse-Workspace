using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public interface IPayrollPoliciesService
    {
        Task<List<EmpAttCategoryModel>> GetCategoriesAsync();
        Task<EmpAttCategoryModel?> GetCategoryByIdAsync(int id);
        Task<bool> SaveCategoryAsync(EmpAttCategoryModel model);
        Task<bool> DeleteCategoryAsync(int id);
    }
}
