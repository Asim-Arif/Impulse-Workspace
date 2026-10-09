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
    public class WorkflowConfigurationDataAccess : IWorkflowConfigurationDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<WorkflowConfigurationDataAccess> _logger;

        public WorkflowConfigurationDataAccess(IConfiguration configuration, ILogger<WorkflowConfigurationDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task<List<WorkflowConfigurationModel>> GetAllWorkflowConfigsAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sqlConfigs = @"
                    SELECT WorkflowCode, WorkflowName, Module, Description, IsEnabled, ApproverRole, UpdatedDate, UpdatedBy
                    FROM dbo.WorkflowConfigurations
                    ORDER BY Module, WorkflowName";

                const string sqlRoles = @"
                    SELECT WorkflowCode, RoleName
                    FROM dbo.WorkflowExemptRoles";

                var configs = (await db.QueryAsync<WorkflowConfigurationModel>(sqlConfigs)).ToList();
                var roles = (await db.QueryAsync<WorkflowExemptRoleModel>(sqlRoles)).ToList();

                foreach (var config in configs)
                {
                    config.ExemptRoles = roles
                        .Where(r => string.Equals(r.WorkflowCode, config.WorkflowCode, StringComparison.OrdinalIgnoreCase))
                        .Select(r => r.RoleName)
                        .ToList();
                }

                return configs;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting all workflow configurations");
                return new List<WorkflowConfigurationModel>();
            }
        }

        public async Task<WorkflowConfigurationModel?> GetWorkflowConfigByCodeAsync(string workflowCode)
        {
            try
            {
                using var db = CreateConnection();
                const string sqlConfig = @"
                    SELECT TOP 1 WorkflowCode, WorkflowName, Module, Description, IsEnabled, ApproverRole, UpdatedDate, UpdatedBy
                    FROM dbo.WorkflowConfigurations
                    WHERE WorkflowCode = @WorkflowCode";

                const string sqlRoles = @"
                    SELECT RoleName
                    FROM dbo.WorkflowExemptRoles
                    WHERE WorkflowCode = @WorkflowCode";

                var config = await db.QueryFirstOrDefaultAsync<WorkflowConfigurationModel>(sqlConfig, new { WorkflowCode = workflowCode });
                if (config != null)
                {
                    var roles = (await db.QueryAsync<string>(sqlRoles, new { WorkflowCode = workflowCode })).ToList();
                    config.ExemptRoles = roles;
                }

                return config;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting workflow configuration for code {WorkflowCode}", workflowCode);
                return null;
            }
        }

        public async Task<bool> UpdateWorkflowConfigAsync(WorkflowConfigurationModel model, string updatedBy)
        {
            try
            {
                using var db = CreateConnection();
                if (db.State != ConnectionState.Open) db.Open();
                using var trans = db.BeginTransaction();

                try
                {
                    const string updateConfigSql = @"
                        UPDATE dbo.WorkflowConfigurations
                        SET IsEnabled = @IsEnabled,
                            ApproverRole = @ApproverRole,
                            UpdatedDate = GETDATE(),
                            UpdatedBy = @UpdatedBy
                        WHERE WorkflowCode = @WorkflowCode";

                    await db.ExecuteAsync(updateConfigSql, new
                    {
                        model.IsEnabled,
                        model.ApproverRole,
                        UpdatedBy = updatedBy,
                        model.WorkflowCode
                    }, trans);

                    const string deleteRolesSql = @"DELETE FROM dbo.WorkflowExemptRoles WHERE WorkflowCode = @WorkflowCode";
                    await db.ExecuteAsync(deleteRolesSql, new { model.WorkflowCode }, trans);

                    if (model.ExemptRoles != null && model.ExemptRoles.Any())
                    {
                        const string insertRoleSql = @"INSERT INTO dbo.WorkflowExemptRoles (WorkflowCode, RoleName) VALUES (@WorkflowCode, @RoleName)";
                        foreach (var role in model.ExemptRoles.Distinct(StringComparer.OrdinalIgnoreCase))
                        {
                            if (!string.IsNullOrWhiteSpace(role))
                            {
                                await db.ExecuteAsync(insertRoleSql, new { model.WorkflowCode, RoleName = role.Trim() }, trans);
                            }
                        }
                    }

                    trans.Commit();
                    return true;
                }
                catch
                {
                    trans.Rollback();
                    throw;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating workflow configuration for code {WorkflowCode}", model.WorkflowCode);
                return false;
            }
        }

        public async Task<bool> IsWorkflowApprovalRequiredAsync(string workflowCode, IEnumerable<string> userRoles)
        {
            try
            {
                var config = await GetWorkflowConfigByCodeAsync(workflowCode);
                if (config == null)
                {
                    // Fallback if not configured: check if user is Director or Admin
                    return !userRoles.Any(r => r.Equals("Director", StringComparison.OrdinalIgnoreCase)
                                            || r.Equals("Admin", StringComparison.OrdinalIgnoreCase)
                                            || r.Equals("Administrator", StringComparison.OrdinalIgnoreCase));
                }

                // If disabled completely, approval is NOT required for anyone (direct save)
                if (!config.IsEnabled)
                {
                    return false;
                }

                // If user possesses any of the exempt roles, approval is NOT required
                if (config.ExemptRoles != null && config.ExemptRoles.Any())
                {
                    bool isExempt = userRoles.Any(userRole =>
                        config.ExemptRoles.Any(exemptRole => string.Equals(exemptRole, userRole, StringComparison.OrdinalIgnoreCase)));

                    if (isExempt)
                    {
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error checking workflow approval requirement for {WorkflowCode}", workflowCode);
                // Fallback safe default
                return !userRoles.Any(r => r.Equals("Director", StringComparison.OrdinalIgnoreCase)
                                        || r.Equals("Admin", StringComparison.OrdinalIgnoreCase)
                                        || r.Equals("Administrator", StringComparison.OrdinalIgnoreCase));
            }
        }
    }
}
