using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Company;
using DataAccessLibrary.Models.Company;
using DataAccessLibrary.Models.ViewModels.Company;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.Company
{
    public class MiscSetupDataAccess : IMiscSetupDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<MiscSetupDataAccess> _logger;

        public MiscSetupDataAccess(IConfiguration configuration, ILogger<MiscSetupDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        public async Task<List<GenericLookupModel>> GetLookupDataAsync(MiscSetupConfig config)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                
                // Using string interpolation for table/column names is safe here because 
                // the config values come from our hardcoded static dictionary, not from user input.
                string sql = $@"
                    SELECT 
                        {config.IdColumn} AS Id, 
                        {config.DescColumn} AS Description 
                    FROM {config.TableName} 
                    ORDER BY {config.DescColumn}";
                    
                var result = await db.QueryAsync<GenericLookupModel>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting lookup data for table {TableName}", config.TableName);
                throw;
            }
        }

        public async Task<bool> InsertLookupDataAsync(MiscSetupConfig config, GenericLookupModel model)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                
                string sql = $@"
                    INSERT INTO {config.TableName} ({config.DescColumn}) 
                    VALUES (@Description)";
                
                var parameters = new { Description = model.Description ?? string.Empty };
                var rowsAffected = await db.ExecuteAsync(sql, parameters);
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error inserting lookup data into table {TableName}", config.TableName);
                throw;
            }
        }

        public async Task<bool> UpdateLookupDataAsync(MiscSetupConfig config, GenericLookupModel model)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                
                string sql = $@"
                    UPDATE {config.TableName} 
                    SET {config.DescColumn} = @Description 
                    WHERE {config.IdColumn} = @Id";
                
                var parameters = new
                {
                    Id = model.Id,
                    Description = model.Description ?? string.Empty
                };

                var rowsAffected = await db.ExecuteAsync(sql, parameters);
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating lookup data in table {TableName} for ID {Id}", config.TableName, model.Id);
                throw;
            }
        }

        public async Task<bool> DeleteLookupDataAsync(MiscSetupConfig config, int id)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                
                string sql = $"DELETE FROM {config.TableName} WHERE {config.IdColumn} = @Id";
                var rowsAffected = await db.ExecuteAsync(sql, new { Id = id });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting lookup data from table {TableName} for ID {Id}", config.TableName, id);
                throw;
            }
        }
    }
}
