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
    public class UserDataAccess : IUserDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<UserDataAccess> _logger;

        public UserDataAccess(IConfiguration configuration, ILogger<UserDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task<List<UserModel>> GetUsersAsync(string? search = null, bool? activeOnly = null)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    SELECT 
                        u.UserID,
                        u.UserName,
                        u.Password,
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
                        COALESCE(u.HelpMainLink, 0) AS HelpMainLink,
                        COALESCE(u.OpenCommandCenter, 0) AS OpenCommandCenter,
                        COALESCE(u.RestrictedItemProfile, 0) AS RestrictedItemProfile,
                        COALESCE(u.GeneralInfoItemProfile, 0) AS GeneralInfoItemProfile,
                        COALESCE(u.ProcessesItemProfile, 0) AS ProcessesItemProfile,
                        COALESCE(u.WeightItemProfile, 0) AS WeightItemProfile,
                        COALESCE(u.PriceItemProfile, 0) AS PriceItemProfile,
                        COALESCE(u.ReferencesItemProfile, 0) AS ReferencesItemProfile,
                        COALESCE(u.PictureItemProfile, 0) AS PictureItemProfile,
                        COALESCE(u.RMItemProfile, 0) AS RMItemProfile,
                        COALESCE(u.ShipInfoItemProfile, 0) AS ShipInfoItemProfile,
                        COALESCE(u.RestrictCompanyCatalogEditing, 0) AS RestrictCompanyCatalogEditing,
                        COALESCE(u.AuthorizeVouchers, 0) AS AuthorizeVouchers,
                        COALESCE(u.AuthorizeIssuance, 0) AS AuthorizeIssuance,
                        COALESCE(u.PostMakerBill, 0) AS PostMakerBill,
                        COALESCE(u.ChangeRateonIssuance, 0) AS ChangeRateonIssuance,
                        COALESCE(u.ChangeRateMakerAssign, 0) AS ChangeRateMakerAssign,
                        COALESCE(u.MakerBill_EditRate, 0) AS MakerBill_EditRate,
                        COALESCE(u.HideRateMakerAssign, 0) AS HideRateMakerAssign,
                        COALESCE(u.AddEditCustomerComplaint, 0) AS AddEditCustomerComplaint,
                        COALESCE(u.EditFollowUp, 0) AS EditFollowUp,
                        COALESCE(u.CloseCAPA, 0) AS CloseCAPA,
                        COALESCE(u.BackupData, 0) AS BackupData,
                        COALESCE(u.RestoreData, 0) AS RestoreData,
                        COALESCE(u.HicoVisible, 0) AS HicoVisible,
                        COALESCE(u.AddProdPlan, 0) AS AddProdPlan,
                        COALESCE(u.ShowTips, 0) AS ShowTips,
                        e.name AS EmployeeName,
                        e.Designation,
                        d.name AS DepartmentName,
                        e.Phone1 AS CellNo
                    FROM Users u
                    LEFT JOIN Employees e ON u.EmpID = e.empid
                    LEFT JOIN Departments d ON e.deptid = d.deptid
                    WHERE (@Search IS NULL OR u.UserName LIKE @SearchParam OR u.FullUserName LIKE @SearchParam OR e.name LIKE @SearchParam OR d.name LIKE @SearchParam)
                      AND (@ActiveOnly IS NULL OR (@ActiveOnly = 1 AND COALESCE(u.InActive, 0) = 0) OR (@ActiveOnly = 0 AND COALESCE(u.InActive, 0) = 1))
                    ORDER BY u.UserName ASC";

                var searchParam = string.IsNullOrWhiteSpace(search) ? null : $"%{search.Trim()}%";
                var result = await db.QueryAsync<UserModel>(sql, new { Search = search, SearchParam = searchParam, ActiveOnly = activeOnly });
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching users list");
                return new List<UserModel>();
            }
        }

        public async Task<UserModel?> GetUserByIdAsync(int userId)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    SELECT 
                        u.UserID,
                        u.UserName,
                        u.Password,
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
                        COALESCE(u.HelpMainLink, 0) AS HelpMainLink,
                        COALESCE(u.OpenCommandCenter, 0) AS OpenCommandCenter,
                        COALESCE(u.RestrictedItemProfile, 0) AS RestrictedItemProfile,
                        COALESCE(u.GeneralInfoItemProfile, 0) AS GeneralInfoItemProfile,
                        COALESCE(u.ProcessesItemProfile, 0) AS ProcessesItemProfile,
                        COALESCE(u.WeightItemProfile, 0) AS WeightItemProfile,
                        COALESCE(u.PriceItemProfile, 0) AS PriceItemProfile,
                        COALESCE(u.ReferencesItemProfile, 0) AS ReferencesItemProfile,
                        COALESCE(u.PictureItemProfile, 0) AS PictureItemProfile,
                        COALESCE(u.RMItemProfile, 0) AS RMItemProfile,
                        COALESCE(u.ShipInfoItemProfile, 0) AS ShipInfoItemProfile,
                        COALESCE(u.RestrictCompanyCatalogEditing, 0) AS RestrictCompanyCatalogEditing,
                        COALESCE(u.AuthorizeVouchers, 0) AS AuthorizeVouchers,
                        COALESCE(u.AuthorizeIssuance, 0) AS AuthorizeIssuance,
                        COALESCE(u.PostMakerBill, 0) AS PostMakerBill,
                        COALESCE(u.ChangeRateonIssuance, 0) AS ChangeRateonIssuance,
                        COALESCE(u.ChangeRateMakerAssign, 0) AS ChangeRateMakerAssign,
                        COALESCE(u.MakerBill_EditRate, 0) AS MakerBill_EditRate,
                        COALESCE(u.HideRateMakerAssign, 0) AS HideRateMakerAssign,
                        COALESCE(u.AddEditCustomerComplaint, 0) AS AddEditCustomerComplaint,
                        COALESCE(u.EditFollowUp, 0) AS EditFollowUp,
                        COALESCE(u.CloseCAPA, 0) AS CloseCAPA,
                        COALESCE(u.BackupData, 0) AS BackupData,
                        COALESCE(u.RestoreData, 0) AS RestoreData,
                        COALESCE(u.HicoVisible, 0) AS HicoVisible,
                        COALESCE(u.AddProdPlan, 0) AS AddProdPlan,
                        COALESCE(u.ShowTips, 0) AS ShowTips,
                        e.name AS EmployeeName,
                        e.Designation,
                        d.name AS DepartmentName,
                        e.Phone1 AS CellNo
                    FROM Users u
                    LEFT JOIN Employees e ON u.EmpID = e.empid
                    LEFT JOIN Departments d ON e.deptid = d.deptid
                    WHERE u.UserID = @UserID";

                return await db.QueryFirstOrDefaultAsync<UserModel>(sql, new { UserID = userId });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching user by ID {UserID}", userId);
                return null;
            }
        }

        public async Task<UserModel?> GetUserByUserNameAsync(string userName)
        {
            try
            {
                using var db = CreateConnection();
                var sql = "SELECT * FROM Users WHERE UserName = @UserName";
                return await db.QueryFirstOrDefaultAsync<UserModel>(sql, new { UserName = userName });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching user by UserName {UserName}", userName);
                return null;
            }
        }

        public async Task<int> CreateUserAsync(UserModel user)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    INSERT INTO Users (
                        UserName, Password, FullUserName, EmpID, InActive,
                        UserManagement, ChangePassword,
                        CompanyMainLink, FinancialMainLink, PayrollMainLink, ExportMainLink, StockMainLink, ProductionMainLink, DashBoardMainLink, QMSMainLink, FixedAssetsMainLink, SamplingMainLink, HelpMainLink, OpenCommandCenter,
                        RestrictedItemProfile, GeneralInfoItemProfile, ProcessesItemProfile, WeightItemProfile, PriceItemProfile, ReferencesItemProfile, PictureItemProfile, RMItemProfile, ShipInfoItemProfile, RestrictCompanyCatalogEditing,
                        AuthorizeVouchers, AuthorizeIssuance, PostMakerBill, ChangeRateonIssuance, ChangeRateMakerAssign, MakerBill_EditRate, HideRateMakerAssign, AddEditCustomerComplaint, EditFollowUp, CloseCAPA,
                        BackupData, RestoreData, HicoVisible, AddProdPlan, ShowTips
                    ) VALUES (
                        @UserName, @Password, @FullUserName, @EmpID, @InActive,
                        @UserManagement, @ChangePassword,
                        @CompanyMainLink, @FinancialMainLink, @PayrollMainLink, @ExportMainLink, @StockMainLink, @ProductionMainLink, @DashBoardMainLink, @QMSMainLink, @FixedAssetsMainLink, @SamplingMainLink, @HelpMainLink, @OpenCommandCenter,
                        @RestrictedItemProfile, @GeneralInfoItemProfile, @ProcessesItemProfile, @WeightItemProfile, @PriceItemProfile, @ReferencesItemProfile, @PictureItemProfile, @RMItemProfile, @ShipInfoItemProfile, @RestrictCompanyCatalogEditing,
                        @AuthorizeVouchers, @AuthorizeIssuance, @PostMakerBill, @ChangeRateonIssuance, @ChangeRateMakerAssign, @MakerBill_EditRate, @HideRateMakerAssign, @AddEditCustomerComplaint, @EditFollowUp, @CloseCAPA,
                        @BackupData, @RestoreData, @HicoVisible, @AddProdPlan, @ShowTips
                    );
                    SELECT CAST(SCOPE_IDENTITY() AS INT);";

                return await db.ExecuteScalarAsync<int>(sql, user);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating user {UserName}", user.UserName);
                throw;
            }
        }

        public async Task<bool> UpdateUserAsync(UserModel user)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    UPDATE Users SET
                        UserName = @UserName,
                        FullUserName = @FullUserName,
                        EmpID = @EmpID,
                        InActive = @InActive,
                        UserManagement = @UserManagement,
                        ChangePassword = @ChangePassword,
                        CompanyMainLink = @CompanyMainLink,
                        FinancialMainLink = @FinancialMainLink,
                        PayrollMainLink = @PayrollMainLink,
                        ExportMainLink = @ExportMainLink,
                        StockMainLink = @StockMainLink,
                        ProductionMainLink = @ProductionMainLink,
                        DashBoardMainLink = @DashBoardMainLink,
                        QMSMainLink = @QMSMainLink,
                        FixedAssetsMainLink = @FixedAssetsMainLink,
                        SamplingMainLink = @SamplingMainLink,
                        HelpMainLink = @HelpMainLink,
                        OpenCommandCenter = @OpenCommandCenter,
                        RestrictedItemProfile = @RestrictedItemProfile,
                        GeneralInfoItemProfile = @GeneralInfoItemProfile,
                        ProcessesItemProfile = @ProcessesItemProfile,
                        WeightItemProfile = @WeightItemProfile,
                        PriceItemProfile = @PriceItemProfile,
                        ReferencesItemProfile = @ReferencesItemProfile,
                        PictureItemProfile = @PictureItemProfile,
                        RMItemProfile = @RMItemProfile,
                        ShipInfoItemProfile = @ShipInfoItemProfile,
                        RestrictCompanyCatalogEditing = @RestrictCompanyCatalogEditing,
                        AuthorizeVouchers = @AuthorizeVouchers,
                        AuthorizeIssuance = @AuthorizeIssuance,
                        PostMakerBill = @PostMakerBill,
                        ChangeRateonIssuance = @ChangeRateonIssuance,
                        ChangeRateMakerAssign = @ChangeRateMakerAssign,
                        MakerBill_EditRate = @MakerBill_EditRate,
                        HideRateMakerAssign = @HideRateMakerAssign,
                        AddEditCustomerComplaint = @AddEditCustomerComplaint,
                        EditFollowUp = @EditFollowUp,
                        CloseCAPA = @CloseCAPA,
                        BackupData = @BackupData,
                        RestoreData = @RestoreData,
                        HicoVisible = @HicoVisible,
                        AddProdPlan = @AddProdPlan,
                        ShowTips = @ShowTips
                    WHERE UserID = @UserID";

                var rows = await db.ExecuteAsync(sql, user);
                return rows > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating user ID {UserID}", user.UserID);
                throw;
            }
        }

        public async Task<bool> ToggleUserStatusAsync(int userId, bool inActive)
        {
            try
            {
                using var db = CreateConnection();
                var sql = "UPDATE Users SET InActive = @InActive WHERE UserID = @UserID";
                var rows = await db.ExecuteAsync(sql, new { UserID = userId, InActive = inActive });
                return rows > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error toggling status for user ID {UserID}", userId);
                return false;
            }
        }

        public async Task<bool> ResetPasswordAsync(int userId, string newPassword)
        {
            try
            {
                using var db = CreateConnection();
                var sql = "UPDATE Users SET Password = @Password WHERE UserID = @UserID";
                var rows = await db.ExecuteAsync(sql, new { UserID = userId, Password = newPassword });
                return rows > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error resetting password for user ID {UserID}", userId);
                return false;
            }
        }

        public async Task<bool> DeleteUserAsync(int userId)
        {
            try
            {
                using var db = CreateConnection();
                var sql = "DELETE FROM Users WHERE UserID = @UserID";
                var rows = await db.ExecuteAsync(sql, new { UserID = userId });
                return rows > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting user ID {UserID}", userId);
                return false;
            }
        }

        public async Task<int> GetTotalUsersCountAsync()
        {
            try
            {
                using var db = CreateConnection();
                var sql = "SELECT COUNT(1) FROM Users";
                return await db.ExecuteScalarAsync<int>(sql);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting total users count");
                return 0;
            }
        }

        public async Task<int> GetActiveUsersCountAsync()
        {
            try
            {
                using var db = CreateConnection();
                var sql = "SELECT COUNT(1) FROM Users WHERE COALESCE(InActive, 0) = 0";
                return await db.ExecuteScalarAsync<int>(sql);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting active users count");
                return 0;
            }
        }

        public async Task<int> CopyUserAsync(int fromUserId, string newUserName, string? password = null, string? fullUserName = null)
        {
            try
            {
                using var db = CreateConnection();
                var parameters = new DynamicParameters();
                parameters.Add("@FromID", fromUserId);
                parameters.Add("@NewUserName", newUserName);

                await db.ExecuteAsync("SP_CopyUser", parameters, commandType: CommandType.StoredProcedure);

                var newId = await db.QueryFirstOrDefaultAsync<int>(
                    "SELECT UserID FROM Users WHERE UserName = @NewUserName",
                    new { NewUserName = newUserName }
                );

                if (newId > 0 && (!string.IsNullOrWhiteSpace(password) || !string.IsNullOrWhiteSpace(fullUserName)))
                {
                    var updates = new List<string>();
                    var updateParams = new DynamicParameters();
                    updateParams.Add("@UserID", newId);

                    if (!string.IsNullOrWhiteSpace(password))
                    {
                        updates.Add("Password = @Password");
                        updateParams.Add("@Password", password);
                    }
                    if (!string.IsNullOrWhiteSpace(fullUserName))
                    {
                        updates.Add("FullUserName = @FullUserName");
                        updateParams.Add("@FullUserName", fullUserName);
                    }

                    var updateSql = $"UPDATE Users SET {string.Join(", ", updates)} WHERE UserID = @UserID";
                    await db.ExecuteAsync(updateSql, updateParams);
                }

                return newId;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error copying user ID {FromID} to {NewUserName}", fromUserId, newUserName);
                throw;
            }
        }
    }
}
