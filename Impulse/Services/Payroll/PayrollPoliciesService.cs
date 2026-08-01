using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.DAC.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public class PayrollPoliciesService : IPayrollPoliciesService
    {
        private readonly IPayrollPoliciesDataAccess _dac;

        public PayrollPoliciesService(IPayrollPoliciesDataAccess dac)
        {
            _dac = dac;
        }

        public async Task<List<EmpAttCategoryModel>> GetCategoriesAsync()
        {
            return await _dac.GetCategoriesAsync();
        }

        public async Task<EmpAttCategoryModel?> GetCategoryByIdAsync(int id)
        {
            return await _dac.GetCategoryByIdAsync(id);
        }

        public async Task<bool> SaveCategoryAsync(EmpAttCategoryModel model)
        {
            return await _dac.SaveCategoryAsync(model);
        }

        public async Task<bool> DeleteCategoryAsync(int id)
        {
            return await _dac.DeleteCategoryAsync(id);
        }
    }
}
