using Dapper;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.Extensions.Configuration;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using System;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.Stock
{
    public class RMGroupsDataAccess : IRMGroupsDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<RMGroupsDataAccess> _logger;

        public RMGroupsDataAccess(IConfiguration configuration, ILogger<RMGroupsDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection") 
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        public async Task<List<RMGroupViewModel>> GetRMGroupsAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "SELECT ID, Description, GrpColor FROM RMGroups ORDER BY Description";
                var result = await db.QueryAsync<RMGroupViewModel>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting RMGroups");
                throw;
            }
        }

        public async Task<int> InsertRMGroupAsync(RMGroupViewModel group)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = @"INSERT INTO RMGroups (Description, GrpColor) 
                               VALUES (@Description, @GrpColor);
                               SELECT CAST(SCOPE_IDENTITY() as int);";
                               
                return await db.QuerySingleAsync<int>(sql, new { group.Description, group.GrpColor });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error inserting RMGroup");
                throw;
            }
        }

        public async Task<int> UpdateRMGroupAsync(RMGroupViewModel group)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = @"UPDATE RMGroups 
                               SET Description = @Description, GrpColor = @GrpColor 
                               WHERE ID = @ID";
                               
                return await db.ExecuteAsync(sql, new { group.Description, group.GrpColor, group.ID });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating RMGroup");
                throw;
            }
        }

        public async Task<int> DeleteRMGroupAsync(int id)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "DELETE FROM RMGroups WHERE ID = @ID";
                return await db.ExecuteAsync(sql, new { ID = id });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting RMGroup");
                throw;
            }
        }

        public async Task<int> CountLinkedMaterialsAsync(int groupId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "SELECT COUNT(*) FROM RM WHERE GroupID = @GroupID";
                return await db.ExecuteScalarAsync<int>(sql, new { GroupID = groupId });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error counting linked materials for RMGroup");
                throw;
            }
        }
    }
}
