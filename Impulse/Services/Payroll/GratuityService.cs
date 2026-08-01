using System.Threading.Tasks;
using DataAccessLibrary.DAC.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public class GratuityService : IGratuityService
    {
        private readonly IGratuityDataAccess _dac;
        private readonly IAuditService _auditService;

        public GratuityService(IGratuityDataAccess dac, IAuditService auditService)
        {
            _dac = dac;
            _auditService = auditService;
        }

        public async Task<GratuitySheetModel> GetGratuitySheetAsync(int year, string? deptId)
        {
            return await _dac.GetGratuitySheetAsync(year, deptId);
        }

        public async Task<bool> SaveGratuitySheetAsync(SaveGratuityDto dto)
        {
            if (string.IsNullOrWhiteSpace(dto.UserName))
            {
                dto.UserName = _auditService.GetCurrentUserName();
            }
            if (string.IsNullOrWhiteSpace(dto.MachineName))
            {
                dto.MachineName = System.Environment.MachineName;
            }

            return await _dac.SaveGratuitySheetAsync(dto);
        }
    }
}
