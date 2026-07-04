using DataAccessLibrary.Models.Company;
using DataAccessLibrary.Models.ViewModels.Company;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Company
{
    public interface IMiscSetupService
    {
        Task<List<GenericLookupModel>> GetLookupDataAsync(MiscSetupConfig config);
        Task<bool> InsertLookupDataAsync(MiscSetupConfig config, GenericLookupModel model);
        Task<bool> UpdateLookupDataAsync(MiscSetupConfig config, GenericLookupModel model);
        Task<bool> DeleteLookupDataAsync(MiscSetupConfig config, int id);
    }
}
