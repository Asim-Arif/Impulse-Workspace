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
    public class UserPermissionDataAccess : IUserPermissionDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<UserPermissionDataAccess> _logger;

        public UserPermissionDataAccess(IConfiguration configuration, ILogger<UserPermissionDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task<List<MenuOptionModel>> GetAllMenuOptionsAsync()
        {
            try
            {
                using var db = CreateConnection();
                var sql = "SELECT OptionID, OptionName, ModuleName, FormName FROM MenuOptions ORDER BY ModuleName, OptionName";
                var list = await db.QueryAsync<MenuOptionModel>(sql);
                return list.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting all menu options");
                return new List<MenuOptionModel>();
            }
        }

        public async Task<List<MenuOptionModel>> GetMenuOptionsByModuleAsync(string moduleName)
        {
            try
            {
                using var db = CreateConnection();
                var sql = "SELECT OptionID, OptionName, ModuleName, FormName FROM MenuOptions WHERE ModuleName = @ModuleName ORDER BY OptionName";
                var list = await db.QueryAsync<MenuOptionModel>(sql, new { ModuleName = moduleName });
                return list.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting menu options for module {Module}", moduleName);
                return new List<MenuOptionModel>();
            }
        }

        public async Task<List<string>> GetDistinctModulesAsync()
        {
            try
            {
                using var db = CreateConnection();
                var sql = "SELECT DISTINCT ModuleName FROM MenuOptions WHERE ModuleName IS NOT NULL AND ModuleName <> '' ORDER BY ModuleName";
                var list = await db.QueryAsync<string>(sql);
                return list.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting distinct modules from MenuOptions");
                return new List<string>();
            }
        }

        public async Task<HashSet<string>> GetUserMenuOptionIdsAsync(int userId)
        {
            try
            {
                using var db = CreateConnection();
                var sql = "SELECT OptionID FROM UserMenuOptions WHERE UserID = @UserId";
                var list = await db.QueryAsync<string>(sql, new { UserId = userId });
                return new HashSet<string>(list, StringComparer.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting user menu options for UserID {UserId}", userId);
                return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }
        }

        public async Task<HashSet<string>> GetUserMenuOptionIdsByUserNameAsync(string userName)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    SELECT umo.OptionID 
                    FROM UserMenuOptions umo 
                    INNER JOIN Users u ON umo.UserID = u.UserID 
                    WHERE u.UserName = @UserName";
                var list = await db.QueryAsync<string>(sql, new { UserName = userName });
                return new HashSet<string>(list, StringComparer.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting user menu options for username {UserName}", userName);
                return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }
        }

        public async Task<UserModel?> GetUserPermissionsByUserNameAsync(string userName)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    SELECT 
                        u.UserID,
                        u.UserName,
                        u.FullUserName,
                        u.EmpID,
                        COALESCE(u.InActive, 0) AS InActive,
                        COALESCE(u.UserManagement, 0) AS UserManagement,
                        COALESCE(u.ChangePassword, 0) AS ChangePassword,
                        COALESCE(u.CompanyMainLink, 0) AS CompanyMainLink,
                        COALESCE(u.FinancialMainLink, 0) AS FinancialMainLink,
                        COALESCE(u.PayrollMainLink, 0) AS PayrollMainLink,
                        COALESCE(u.ExportMainLink, 0) AS ExportMainLink,
                        COALESCE(u.StockMainLink, 0) AS StockMainLink,
                        COALESCE(u.ProductionMainLink, 0) AS ProductionMainLink,
                        COALESCE(u.DashBoardMainLink, 0) AS DashBoardMainLink,
                        COALESCE(u.QMSMainLink, 0) AS QMSMainLink,
                        COALESCE(u.FixedAssetsMainLink, 0) AS FixedAssetsMainLink,
                        COALESCE(u.SamplingMainLink, 0) AS SamplingMainLink,
                        COALESCE(u.HelpMainLink, 0) AS HelpMainLink
                    FROM Users u
                    WHERE u.UserName = @UserName";

                return await db.QueryFirstOrDefaultAsync<UserModel>(sql, new { UserName = userName });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting user permissions model for username {UserName}", userName);
                return null;
            }
        }

        public async Task<bool> SaveUserMenuOptionsAsync(int userId, string moduleName, IEnumerable<string> selectedOptionIds)
        {
            try
            {
                using var db = CreateConnection();
                db.Open();
                using var trans = db.BeginTransaction();

                // 1. Delete existing options for this specific module
                var deleteSql = @"
                    DELETE umo 
                    FROM UserMenuOptions umo
                    INNER JOIN MenuOptions mo ON umo.OptionID = mo.OptionID
                    WHERE umo.UserID = @UserId AND mo.ModuleName = @ModuleName";

                await db.ExecuteAsync(deleteSql, new { UserId = userId, ModuleName = moduleName }, trans);

                // 2. Insert selected options
                if (selectedOptionIds != null && selectedOptionIds.Any())
                {
                    var insertSql = "INSERT INTO UserMenuOptions (UserID, OptionID) VALUES (@UserId, @OptionID)";
                    var items = selectedOptionIds.Select(opt => new { UserId = userId, OptionID = opt });
                    await db.ExecuteAsync(insertSql, items, trans);
                }

                trans.Commit();
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving user menu options for UserID {UserId}, Module {Module}", userId, moduleName);
                return false;
            }
        }

        public async Task<bool> SaveAllUserMenuOptionsAsync(int userId, IEnumerable<string> allOptionIds)
        {
            try
            {
                using var db = CreateConnection();
                db.Open();
                using var trans = db.BeginTransaction();

                // 1. Delete all existing options for this user
                var deleteSql = "DELETE FROM UserMenuOptions WHERE UserID = @UserId";
                await db.ExecuteAsync(deleteSql, new { UserId = userId }, trans);

                // 2. Insert selected options
                if (allOptionIds != null && allOptionIds.Any())
                {
                    var insertSql = "INSERT INTO UserMenuOptions (UserID, OptionID) VALUES (@UserId, @OptionID)";
                    var items = allOptionIds.Select(opt => new { UserId = userId, OptionID = opt });
                    await db.ExecuteAsync(insertSql, items, trans);
                }

                trans.Commit();
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving all user menu options for UserID {UserId}", userId);
                return false;
            }
        }

        public async Task<bool> CopyUserMenuOptionsAsync(int fromUserId, int toUserId)
        {
            try
            {
                using var db = CreateConnection();
                db.Open();
                using var trans = db.BeginTransaction();

                await db.ExecuteAsync("DELETE FROM UserMenuOptions WHERE UserID = @ToUserId", new { ToUserId = toUserId }, trans);

                var copySql = @"
                    INSERT INTO UserMenuOptions (UserID, OptionID)
                    SELECT @ToUserId, OptionID FROM UserMenuOptions WHERE UserID = @FromUserId";

                await db.ExecuteAsync(copySql, new { FromUserId = fromUserId, ToUserId = toUserId }, trans);

                trans.Commit();
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error copying user menu options from {From} to {To}", fromUserId, toUserId);
                return false;
            }
        }

        public async Task<bool> SyncMenuOptionsAsync(IEnumerable<MenuOptionModel> options)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = @OptionID)
                    BEGIN
                        INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName)
                        VALUES (@OptionID, @OptionName, @ModuleName, @FormName)
                    END
                    ELSE
                    BEGIN
                        UPDATE MenuOptions 
                        SET OptionName = @OptionName, ModuleName = @ModuleName, FormName = @FormName
                        WHERE OptionID = @OptionID
                    END";

                foreach (var opt in options)
                {
                    await db.ExecuteAsync(sql, opt);
                }
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error syncing menu options");
                return false;
            }
        }
    }
}
