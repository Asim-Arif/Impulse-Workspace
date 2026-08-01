using System.Threading.Tasks;
using DataAccessLibrary.DAC.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public class SocialSecurityService : ISocialSecurityService
    {
        private readonly ISocialSecurityDataAccess _dac;

        public SocialSecurityService(ISocialSecurityDataAccess dac)
        {
            _dac = dac;
        }

        public async Task<SocialSecuritySheetModel> GetSocialSecuritySheetAsync(SocialSecurityFilterDto filter)
        {
            return await _dac.GetSocialSecuritySheetAsync(filter);
        }

        public async Task<SocialSecuritySheetModel> GenerateSocialSecuritySheetAsync(SocialSecurityFilterDto filter)
        {
            return await _dac.GenerateSocialSecuritySheetAsync(filter);
        }

        public async Task<bool> SaveSocialSecuritySheetAsync(SaveSocialSecurityDto dto)
        {
            return await _dac.SaveSocialSecuritySheetAsync(dto);
        }
    }
}
