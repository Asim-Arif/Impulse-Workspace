using System.Threading.Tasks;
using DataAccessLibrary.DAC.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public class AbsentSheetService : IAbsentSheetService
    {
        private readonly IAbsentSheetDataAccess _dac;

        public AbsentSheetService(IAbsentSheetDataAccess dac)
        {
            _dac = dac;
        }

        public async Task<AbsentSheetResultModel> GetAbsentSheetAsync(AbsentSheetFilterDto filter)
        {
            return await _dac.GetAbsentSheetAsync(filter);
        }
    }
}
