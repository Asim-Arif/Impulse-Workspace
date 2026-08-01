using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.DAC.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public class DesignationService : IDesignationService
    {
        private readonly IDesignationDataAccess _dac;

        public DesignationService(IDesignationDataAccess dac)
        {
            _dac = dac;
        }

        public async Task<List<DesignationModel>> GetDesignationsAsync()
        {
            return await _dac.GetDesignationsAsync();
        }

        public async Task<DesignationModel?> GetDesignationByIdAsync(int entryId)
        {
            return await _dac.GetDesignationByIdAsync(entryId);
        }

        public async Task<bool> SaveDesignationAsync(DesignationModel model)
        {
            return await _dac.SaveDesignationAsync(model);
        }

        public async Task<bool> DeleteDesignationAsync(int entryId)
        {
            return await _dac.DeleteDesignationAsync(entryId);
        }
    }
}
