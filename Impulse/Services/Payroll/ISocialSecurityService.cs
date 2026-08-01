using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public interface ISocialSecurityService
    {
        Task<SocialSecuritySheetModel> GetSocialSecuritySheetAsync(SocialSecurityFilterDto filter);
        Task<SocialSecuritySheetModel> GenerateSocialSecuritySheetAsync(SocialSecurityFilterDto filter);
        Task<bool> SaveSocialSecuritySheetAsync(SaveSocialSecurityDto dto);
    }
}
