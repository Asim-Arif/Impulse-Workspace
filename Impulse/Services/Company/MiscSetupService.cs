using DataAccessLibrary.Interface.Company;
using DataAccessLibrary.Models.Company;
using DataAccessLibrary.Models.ViewModels.Company;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Company
{
    public class MiscSetupService : IMiscSetupService
    {
        private readonly IMiscSetupDataAccess _dataAccess;

        public MiscSetupService(IMiscSetupDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<GenericLookupModel>> GetLookupDataAsync(MiscSetupConfig config)
        {
            return _dataAccess.GetLookupDataAsync(config);
        }

        public Task<bool> InsertLookupDataAsync(MiscSetupConfig config, GenericLookupModel model)
        {
            return _dataAccess.InsertLookupDataAsync(config, model);
        }

        public Task<bool> UpdateLookupDataAsync(MiscSetupConfig config, GenericLookupModel model)
        {
            return _dataAccess.UpdateLookupDataAsync(config, model);
        }

        public Task<bool> DeleteLookupDataAsync(MiscSetupConfig config, int id)
        {
            return _dataAccess.DeleteLookupDataAsync(config, id);
        }
    }
}
