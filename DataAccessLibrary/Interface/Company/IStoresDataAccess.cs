using DataAccessLibrary.Models.ViewModels.Company;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Company
{
    public interface IStoresDataAccess
    {
        Task<List<StoreNodeModel>> GetFullTreeAsync();
        
        Task<bool> InsertStoreAsync(string storeName);
        Task<bool> UpdateStoreAsync(int storeId, string storeName);
        Task<bool> DeleteStoreAsync(int storeId);

        Task<bool> InsertRackAsync(int storeId, string rackNo);
        Task<bool> UpdateRackAsync(int rackId, string rackNo);
        Task<bool> DeleteRackAsync(int rackId);

        Task<bool> InsertBinAsync(int rackId, string shelfNo);
        Task<bool> UpdateBinAsync(int binId, string shelfNo);
        Task<bool> DeleteBinAsync(int binId);
    }
}
