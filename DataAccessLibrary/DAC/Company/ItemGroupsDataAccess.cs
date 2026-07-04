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
    public class ItemGroupsDataAccess : IItemGroupsDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<ItemGroupsDataAccess> _logger;

        public ItemGroupsDataAccess(IConfiguration configuration, ILogger<ItemGroupsDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        public async Task<List<ItemGroupsViewModel>> GetItemGroupsAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = @"
                    SELECT 
                        ID, 
                        Description, 
                        GrpColor, 
                        MainGroup_RefID, 
                        EmpID_Purchaser, 
                        name AS PurchaserName 
                    FROM VItemGroups 
                    ORDER BY Description";
                var result = await db.QueryAsync<ItemGroupsViewModel>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting item groups");
                throw;
            }
        }

        public async Task<List<PurchaserLookupModel>> GetPurchasersAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "SELECT EmpID, Name FROM VEmp WHERE Active = 1 ORDER BY Name";
                var result = await db.QueryAsync<PurchaserLookupModel>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting purchasers");
                throw;
            }
        }

        public async Task<bool> InsertItemGroupAsync(ItemGroupsViewModel model)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = @"
                    INSERT INTO ItemGroups (Description, GrpColor, EmpID_Purchaser) 
                    VALUES (@Description, @GrpColor, @EmpID_Purchaser)";
                
                var parameters = new
                {
                    Description = model.Description ?? string.Empty,
                    GrpColor = model.GrpColor,
                    EmpID_Purchaser = model.EmpID_Purchaser ?? string.Empty
                };

                var rowsAffected = await db.ExecuteAsync(sql, parameters);
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error inserting item group");
                throw;
            }
        }

        public async Task<bool> UpdateItemGroupAsync(ItemGroupsViewModel model)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = @"
                    UPDATE ItemGroups 
                    SET Description = @Description, 
                        GrpColor = @GrpColor, 
                        EmpID_Purchaser = @EmpID_Purchaser 
                    WHERE ID = @ID";
                
                var parameters = new
                {
                    ID = model.ID,
                    Description = model.Description ?? string.Empty,
                    GrpColor = model.GrpColor,
                    EmpID_Purchaser = model.EmpID_Purchaser ?? string.Empty
                };

                var rowsAffected = await db.ExecuteAsync(sql, parameters);
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating item group {ID}", model.ID);
                throw;
            }
        }

        public async Task<bool> DeleteItemGroupAsync(int id)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                const string sql = "DELETE FROM ItemGroups WHERE ID = @ID";
                var rowsAffected = await db.ExecuteAsync(sql, new { ID = id });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting item group {ID}", id);
                throw;
            }
        }
    }
}
