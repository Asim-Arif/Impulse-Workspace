using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Setup;
using DataAccessLibrary.Models.Setup;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.Setup
{
    public class SystemFieldConfigurationDataAccess : ISystemFieldConfigurationDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<SystemFieldConfigurationDataAccess> _logger;

        public SystemFieldConfigurationDataAccess(IConfiguration configuration, ILogger<SystemFieldConfigurationDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task<List<SystemFieldConfigurationModel>> GetConfigurationsByFormAsync(string module, string formKey)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT SettingID, Module, FormKey, FieldKey, IsRequired, IsVisible, Description
                    FROM dbo.SystemFieldConfigurations
                    WHERE Module = @Module AND FormKey = @FormKey";
                var result = await db.QueryAsync<SystemFieldConfigurationModel>(sql, new { Module = module, FormKey = formKey });
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting SystemFieldConfigurations for Module={Module}, FormKey={FormKey}", module, formKey);
                return new List<SystemFieldConfigurationModel>();
            }
        }

        public async Task<List<SystemFieldConfigurationModel>> GetAllConfigurationsAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT SettingID, Module, FormKey, FieldKey, IsRequired, IsVisible, Description
                    FROM dbo.SystemFieldConfigurations
                    ORDER BY Module, FormKey, FieldKey";
                var result = await db.QueryAsync<SystemFieldConfigurationModel>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting all SystemFieldConfigurations");
                return new List<SystemFieldConfigurationModel>();
            }
        }

        public async Task<bool> UpdateConfigurationAsync(int settingId, bool isRequired, bool isVisible)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    UPDATE dbo.SystemFieldConfigurations
                    SET IsRequired = @IsRequired, IsVisible = @IsVisible
                    WHERE SettingID = @SettingID";
                int affected = await db.ExecuteAsync(sql, new { SettingID = settingId, IsRequired = isRequired, IsVisible = isVisible });
                return affected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating SystemFieldConfiguration SettingID={SettingID}", settingId);
                return false;
            }
        }
    }
}

