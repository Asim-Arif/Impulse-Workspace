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
    public class UserRoleDataAccess : IUserRoleDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<UserRoleDataAccess> _logger;

        public UserRoleDataAccess(IConfiguration configuration, ILogger<UserRoleDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task<List<UserRoleModel>> GetAllRolesAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT r.User_Role, COUNT(ur.UserID) AS AssignedUsersCount
                    FROM User_Roles r
                    LEFT JOIN Users_User_Roles ur ON r.User_Role = ur.User_Role
                    GROUP BY r.User_Role
                    ORDER BY r.User_Role";

                var result = await db.QueryAsync<UserRoleModel>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching all user roles");
                return new List<UserRoleModel>();
            }
        }

        public async Task<bool> AddRoleAsync(string roleName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(roleName)) return false;
                var trimmed = roleName.Trim();

                using var db = CreateConnection();
                const string sql = @"
                    IF NOT EXISTS (SELECT 1 FROM User_Roles WHERE User_Role = @RoleName)
                    BEGIN
                        INSERT INTO User_Roles (User_Role) VALUES (@RoleName);
                        SELECT 1;
                    END
                    ELSE
                    BEGIN
                        SELECT 0;
                    END";

                var rows = await db.ExecuteScalarAsync<int>(sql, new { RoleName = trimmed });
                return rows > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error adding user role {RoleName}", roleName);
                return false;
            }
        }

        public async Task<bool> DeleteRoleAsync(string roleName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(roleName)) return false;
                var trimmed = roleName.Trim();

                using var db = CreateConnection();
                // Cascade constraints will remove associations from Users_User_Roles
                const string sql = "DELETE FROM User_Roles WHERE User_Role = @RoleName";
                var rows = await db.ExecuteAsync(sql, new { RoleName = trimmed });
                return rows > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting user role {RoleName}", roleName);
                return false;
            }
        }

        public async Task<bool> RoleExistsAsync(string roleName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(roleName)) return false;
                var trimmed = roleName.Trim();

                using var db = CreateConnection();
                const string sql = "SELECT COUNT(1) FROM User_Roles WHERE User_Role = @RoleName";
                var count = await db.ExecuteScalarAsync<int>(sql, new { RoleName = trimmed });
                return count > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error checking role existence for {RoleName}", roleName);
                return false;
            }
        }

        public async Task<List<string>> GetRolesByUserIdAsync(int userId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "SELECT User_Role FROM Users_User_Roles WHERE UserID = @UserID ORDER BY User_Role";
                var result = await db.QueryAsync<string>(sql, new { UserID = userId });
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting roles for UserID {UserID}", userId);
                return new List<string>();
            }
        }

        public async Task<Dictionary<int, List<string>>> GetAllUserRolesMapAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "SELECT UserID, User_Role FROM Users_User_Roles ORDER BY UserID, User_Role";
                var rows = await db.QueryAsync<(int UserID, string User_Role)>(sql);

                var map = new Dictionary<int, List<string>>();
                foreach (var (uId, role) in rows)
                {
                    if (!map.TryGetValue(uId, out var list))
                    {
                        list = new List<string>();
                        map[uId] = list;
                    }
                    list.Add(role);
                }
                return map;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting all user roles map");
                return new Dictionary<int, List<string>>();
            }
        }

        public async Task<List<UserModel>> GetUsersByRoleAsync(string roleName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(roleName)) return new List<UserModel>();
                var trimmed = roleName.Trim();

                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        u.UserID,
                        u.UserName,
                        u.FullUserName,
                        u.EmpID,
                        COALESCE(u.InActive, 0) AS InActive,
                        e.Name AS EmployeeName,
                        e.Designation,
                        d.name AS DepartmentName,
                        e.Phone1 AS CellNo
                    FROM Users u
                    INNER JOIN Users_User_Roles ur ON u.UserID = ur.UserID
                    LEFT JOIN Employees e ON u.EmpID = e.EmpID
                    LEFT JOIN Departments d ON e.deptid = d.deptid
                    WHERE ur.User_Role = @RoleName AND COALESCE(u.InActive, 0) = 0
                    ORDER BY u.UserName";

                var result = await db.QueryAsync<UserModel>(sql, new { RoleName = trimmed });
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting users by role {RoleName}", roleName);
                return new List<UserModel>();
            }
        }

        public async Task<bool> SaveUserRolesAsync(int userId, IEnumerable<string> roles)
        {
            if (userId <= 0) return false;

            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            using var trans = conn.BeginTransaction();

            try
            {
                const string deleteSql = "DELETE FROM Users_User_Roles WHERE UserID = @UserID";
                await conn.ExecuteAsync(deleteSql, new { UserID = userId }, trans);

                var distinctRoles = roles?
                    .Where(r => !string.IsNullOrWhiteSpace(r))
                    .Select(r => r.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (distinctRoles != null && distinctRoles.Count > 0)
                {
                    const string insertSql = "INSERT INTO Users_User_Roles (UserID, User_Role) VALUES (@UserID, @User_Role)";
                    var items = distinctRoles.Select(r => new { UserID = userId, User_Role = r });
                    await conn.ExecuteAsync(insertSql, items, trans);
                }

                trans.Commit();
                return true;
            }
            catch (Exception ex)
            {
                trans.Rollback();
                _logger.LogError(ex, "Error saving roles for UserID {UserID}", userId);
                return false;
            }
        }

        public async Task<bool> CopyUserRolesAsync(int sourceUserId, int targetUserId)
        {
            if (sourceUserId <= 0 || targetUserId <= 0) return false;

            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    DELETE FROM Users_User_Roles WHERE UserID = @TargetUserId;
                    INSERT INTO Users_User_Roles (UserID, User_Role)
                    SELECT @TargetUserId, User_Role
                    FROM Users_User_Roles
                    WHERE UserID = @SourceUserId;";

                await db.ExecuteAsync(sql, new { SourceUserId = sourceUserId, TargetUserId = targetUserId });
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error copying roles from {SourceUserId} to {TargetUserId}", sourceUserId, targetUserId);
                return false;
            }
        }
    }
}
