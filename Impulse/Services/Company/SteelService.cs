using DataAccessLibrary.Interface.Company;
using DataAccessLibrary.Models.ViewModels.Company;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Company
{
    public class SteelService : ISteelService
    {
        private readonly ISteelDataAccess _dataAccess;

        public SteelService(ISteelDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<SteelTypeNodeModel>> GetFullSteelTreeAsync()
        {
            return _dataAccess.GetFullSteelTreeAsync();
        }

        public Task<List<SteelCompModel>> GetAllCompositionsAsync()
        {
            return _dataAccess.GetAllCompositionsAsync();
        }

        public Task<List<string>> GetUnitsAsync()
        {
            return _dataAccess.GetUnitsAsync();
        }

        public Task<bool> SaveSteelTypeAsync(SteelTypeNodeModel model)
        {
            return _dataAccess.SaveSteelTypeAsync(model);
        }

        public Task<bool> DeleteSteelTypeAsync(int steelId)
        {
            return _dataAccess.DeleteSteelTypeAsync(steelId);
        }

        public Task<bool> SaveSteelGageAsync(SteelGageNodeModel model, bool isInsert)
        {
            return _dataAccess.SaveSteelGageAsync(model, isInsert);
        }

        public Task<bool> DeleteSteelGageAsync(string gageId)
        {
            return _dataAccess.DeleteSteelGageAsync(gageId);
        }
    }
}
