using DataAccessLibrary.Interface.Company;
using DataAccessLibrary.Models.ViewModels.Company;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Company
{
    public class StoresService : IStoresService
    {
        private readonly IStoresDataAccess _dataAccess;

        public StoresService(IStoresDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<StoreNodeModel>> GetFullTreeAsync()
        {
            return _dataAccess.GetFullTreeAsync();
        }

        public Task<bool> InsertStoreAsync(string storeName)
        {
            return _dataAccess.InsertStoreAsync(storeName);
        }

        public Task<bool> UpdateStoreAsync(int storeId, string storeName)
        {
            return _dataAccess.UpdateStoreAsync(storeId, storeName);
        }

        public Task<bool> DeleteStoreAsync(int storeId)
        {
            return _dataAccess.DeleteStoreAsync(storeId);
        }

        public Task<bool> InsertRackAsync(int storeId, string rackNo)
        {
            return _dataAccess.InsertRackAsync(storeId, rackNo);
        }

        public Task<bool> UpdateRackAsync(int rackId, string rackNo)
        {
            return _dataAccess.UpdateRackAsync(rackId, rackNo);
        }

        public Task<bool> DeleteRackAsync(int rackId)
        {
            return _dataAccess.DeleteRackAsync(rackId);
        }

        public Task<bool> InsertBinAsync(int rackId, string shelfNo)
        {
            return _dataAccess.InsertBinAsync(rackId, shelfNo);
        }

        public Task<bool> UpdateBinAsync(int binId, string shelfNo)
        {
            return _dataAccess.UpdateBinAsync(binId, shelfNo);
        }

        public Task<bool> DeleteBinAsync(int binId)
        {
            return _dataAccess.DeleteBinAsync(binId);
        }
    }
}
