using System.Threading.Tasks;
using DataAccessLibrary.DAC.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public class EOBIService : IEOBIService
    {
        private readonly IEOBIDataAccess _dac;

        public EOBIService(IEOBIDataAccess dac)
        {
            _dac = dac;
        }

        public async Task<EOBISheetModel> GetEOBISheetAsync(EOBIFilterDto filter)
        {
            return await _dac.GetEOBISheetAsync(filter);
        }

        public async Task<EOBISheetModel> GenerateEOBISheetAsync(EOBIFilterDto filter)
        {
            return await _dac.GenerateEOBISheetAsync(filter);
        }

        public async Task<bool> SaveEOBISheetAsync(SaveEOBIDto dto)
        {
            return await _dac.SaveEOBISheetAsync(dto);
        }
    }
}
