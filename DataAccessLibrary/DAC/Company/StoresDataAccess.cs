using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Company;
using DataAccessLibrary.Models.ViewModels.Company;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.Company
{
    public class StoresDataAccess : IStoresDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<StoresDataAccess> _logger;

        public StoresDataAccess(IConfiguration configuration, ILogger<StoresDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        public async Task<List<StoreNodeModel>> GetFullTreeAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                
                var stores = await db.QueryAsync<StoreNodeModel>(
                    "SELECT EntryID as StoreID, StoreName FROM Stores ORDER BY StoreName");
                    
                var racks = await db.QueryAsync<RackNodeModel>(
                    "SELECT EntryID as RackID, Store_RefID as StoreID, RackNo FROM StoreRacks ORDER BY RackNo");
                    
                var bins = await db.QueryAsync<BinNodeModel>(
                    "SELECT EntryID as BinID, Rack_RefID as RackID, ShelfNo FROM StoreShelfs ORDER BY ShelfNo");

                var storeList = stores.ToList();
                var rackList = racks.ToList();
                var binList = bins.ToList();

                foreach (var rack in rackList)
                {
                    rack.Bins = binList.Where(b => b.RackID == rack.RackID).ToList();
                }

                foreach (var store in storeList)
                {
                    store.Racks = rackList.Where(r => r.StoreID == store.StoreID).ToList();
                }

                return storeList;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting stores tree.");
                throw;
            }
        }

        public async Task<bool> InsertStoreAsync(string storeName)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "INSERT INTO Stores (StoreName) VALUES (@StoreName)";
                var rowsAffected = await db.ExecuteAsync(sql, new { StoreName = storeName });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error inserting store.");
                throw;
            }
        }

        public async Task<bool> UpdateStoreAsync(int storeId, string storeName)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "UPDATE Stores SET StoreName = @StoreName WHERE EntryID = @StoreId";
                var rowsAffected = await db.ExecuteAsync(sql, new { StoreName = storeName, StoreId = storeId });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating store.");
                throw;
            }
        }

        public async Task<bool> DeleteStoreAsync(int storeId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "DELETE FROM Stores WHERE EntryID = @StoreId";
                var rowsAffected = await db.ExecuteAsync(sql, new { StoreId = storeId });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting store.");
                throw;
            }
        }

        public async Task<bool> InsertRackAsync(int storeId, string rackNo)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "INSERT INTO StoreRacks (Store_RefID, RackNo) VALUES (@StoreId, @RackNo)";
                var rowsAffected = await db.ExecuteAsync(sql, new { StoreId = storeId, RackNo = rackNo });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error inserting rack.");
                throw;
            }
        }

        public async Task<bool> UpdateRackAsync(int rackId, string rackNo)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "UPDATE StoreRacks SET RackNo = @RackNo WHERE EntryID = @RackId";
                var rowsAffected = await db.ExecuteAsync(sql, new { RackNo = rackNo, RackId = rackId });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating rack.");
                throw;
            }
        }

        public async Task<bool> DeleteRackAsync(int rackId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "DELETE FROM StoreRacks WHERE EntryID = @RackId";
                var rowsAffected = await db.ExecuteAsync(sql, new { RackId = rackId });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting rack.");
                throw;
            }
        }

        public async Task<bool> InsertBinAsync(int rackId, string shelfNo)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "INSERT INTO StoreShelfs (Rack_RefID, ShelfNo) VALUES (@RackId, @ShelfNo)";
                var rowsAffected = await db.ExecuteAsync(sql, new { RackId = rackId, ShelfNo = shelfNo });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error inserting bin.");
                throw;
            }
        }

        public async Task<bool> UpdateBinAsync(int binId, string shelfNo)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "UPDATE StoreShelfs SET ShelfNo = @ShelfNo WHERE EntryID = @BinId";
                var rowsAffected = await db.ExecuteAsync(sql, new { ShelfNo = shelfNo, BinId = binId });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating bin.");
                throw;
            }
        }

        public async Task<bool> DeleteBinAsync(int binId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "DELETE FROM StoreShelfs WHERE EntryID = @BinId";
                var rowsAffected = await db.ExecuteAsync(sql, new { BinId = binId });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting bin.");
                throw;
            }
        }
    }
}
