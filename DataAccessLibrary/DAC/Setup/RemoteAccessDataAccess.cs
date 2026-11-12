using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Setup;
using DataAccessLibrary.Models.Setup;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.Setup
{
    public class RemoteAccessDataAccess : IRemoteAccessDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<RemoteAccessDataAccess> _logger;

        public RemoteAccessDataAccess(IConfiguration configuration, ILogger<RemoteAccessDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task<int> CreateRequestAsync(RemoteAccessRequestModel request)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    INSERT INTO RemoteAccessRequests (
                        UserID, UserName, ClientIP, RequestDT, Reason, Status
                    ) VALUES (
                        @UserID, @UserName, @ClientIP, GETDATE(), 'Pending', 'Pending'
                    );
                    SELECT CAST(SCOPE_IDENTITY() AS INT);";

                // Ensure Reason is populated
                var param = new
                {
                    request.UserID,
                    request.UserName,
                    request.ClientIP,
                    Reason = string.IsNullOrWhiteSpace(request.Reason) ? "Remote working authorization requested" : request.Reason.Trim()
                };

                const string insertSql = @"
                    INSERT INTO RemoteAccessRequests (
                        UserID, UserName, ClientIP, RequestDT, Reason, Status
                    ) VALUES (
                        @UserID, @UserName, @ClientIP, GETDATE(), @Reason, 'Pending'
                    );
                    SELECT CAST(SCOPE_IDENTITY() AS INT);";

                return await db.ExecuteScalarAsync<int>(insertSql, param);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating remote access request for user {UserName} from IP {ClientIP}", request.UserName, request.ClientIP);
                throw;
            }
        }

        public async Task<RemoteAccessRequestModel?> GetActiveApprovalAsync(int userId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT TOP 1 
                        r.EntryID, r.UserID, r.UserName, r.ClientIP, r.RequestDT, r.Reason, r.Status,
                        r.DurationMinutes, r.ExpiresAt, r.ActionDT, r.ActionByUserID, r.ActionByUserName, r.Remarks,
                        e.name AS EmployeeName,
                        d.name AS DepartmentName,
                        e.Designation
                    FROM RemoteAccessRequests r
                    INNER JOIN Users u ON r.UserID = u.UserID
                    LEFT JOIN Employees e ON u.EmpID = e.empid
                    LEFT JOIN Departments d ON e.deptid = d.deptid
                    WHERE r.UserID = @UserID 
                      AND r.Status = 'Approved' 
                      AND r.ExpiresAt > GETDATE()
                    ORDER BY r.ExpiresAt DESC";

                return await db.QueryFirstOrDefaultAsync<RemoteAccessRequestModel>(sql, new { UserID = userId });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error checking active remote approval for user ID {UserID}", userId);
                return null;
            }
        }

        public async Task<RemoteAccessRequestModel?> GetLatestPendingRequestAsync(int userId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT TOP 1 
                        r.EntryID, r.UserID, r.UserName, r.ClientIP, r.RequestDT, r.Reason, r.Status,
                        r.DurationMinutes, r.ExpiresAt, r.ActionDT, r.ActionByUserID, r.ActionByUserName, r.Remarks
                    FROM RemoteAccessRequests r
                    WHERE r.UserID = @UserID 
                      AND r.Status = 'Pending'
                      AND r.RequestDT > DATEADD(MINUTE, -30, GETDATE())
                    ORDER BY r.RequestDT DESC";

                return await db.QueryFirstOrDefaultAsync<RemoteAccessRequestModel>(sql, new { UserID = userId });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting latest pending request for user ID {UserID}", userId);
                return null;
            }
        }

        public async Task<RemoteAccessRequestModel?> GetLatestRequestByUserAsync(int userId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT TOP 1 
                        r.EntryID, r.UserID, r.UserName, r.ClientIP, r.RequestDT, r.Reason, r.Status,
                        r.DurationMinutes, r.ExpiresAt, r.ActionDT, r.ActionByUserID, r.ActionByUserName, r.Remarks,
                        e.name AS EmployeeName,
                        d.name AS DepartmentName,
                        e.Designation
                    FROM RemoteAccessRequests r
                    INNER JOIN Users u ON r.UserID = u.UserID
                    LEFT JOIN Employees e ON u.EmpID = e.empid
                    LEFT JOIN Departments d ON e.deptid = d.deptid
                    WHERE r.UserID = @UserID
                    ORDER BY r.EntryID DESC";

                return await db.QueryFirstOrDefaultAsync<RemoteAccessRequestModel>(sql, new { UserID = userId });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting latest request for UserID {UserID}", userId);
                return null;
            }
        }

        public async Task<List<RemoteAccessRequestModel>> GetPendingRequestsAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        r.EntryID, r.UserID, r.UserName, r.ClientIP, r.RequestDT, r.Reason, r.Status,
                        r.DurationMinutes, r.ExpiresAt, r.ActionDT, r.ActionByUserID, r.ActionByUserName, r.Remarks,
                        e.name AS EmployeeName,
                        d.name AS DepartmentName,
                        e.Designation
                    FROM RemoteAccessRequests r
                    INNER JOIN Users u ON r.UserID = u.UserID
                    LEFT JOIN Employees e ON u.EmpID = e.empid
                    LEFT JOIN Departments d ON e.deptid = d.deptid
                    WHERE r.Status = 'Pending'
                    ORDER BY r.RequestDT ASC";

                var results = await db.QueryAsync<RemoteAccessRequestModel>(sql);
                return results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching pending remote access requests");
                return new List<RemoteAccessRequestModel>();
            }
        }

        public async Task<List<RemoteAccessRequestModel>> GetRecentRequestsAsync(int limit = 50)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT TOP (@Limit)
                        r.EntryID, r.UserID, r.UserName, r.ClientIP, r.RequestDT, r.Reason, r.Status,
                        r.DurationMinutes, r.ExpiresAt, r.ActionDT, r.ActionByUserID, r.ActionByUserName, r.Remarks,
                        e.name AS EmployeeName,
                        d.name AS DepartmentName,
                        e.Designation
                    FROM RemoteAccessRequests r
                    INNER JOIN Users u ON r.UserID = u.UserID
                    LEFT JOIN Employees e ON u.EmpID = e.empid
                    LEFT JOIN Departments d ON e.deptid = d.deptid
                    ORDER BY r.RequestDT DESC";

                var results = await db.QueryAsync<RemoteAccessRequestModel>(sql, new { Limit = limit });
                return results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching recent remote access requests");
                return new List<RemoteAccessRequestModel>();
            }
        }

        public async Task<bool> ApproveRequestAsync(int entryId, int actionByUserId, string actionByUserName, int durationMinutes, string? remarks)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    UPDATE RemoteAccessRequests 
                    SET Status = 'Approved',
                        DurationMinutes = @DurationMinutes,
                        ExpiresAt = DATEADD(MINUTE, @DurationMinutes, GETDATE()),
                        ActionDT = GETDATE(),
                        ActionByUserID = @ActionByUserId,
                        ActionByUserName = @ActionByUserName,
                        Remarks = @Remarks
                    WHERE EntryID = @EntryID AND Status = 'Pending';";

                var affected = await db.ExecuteAsync(sql, new
                {
                    EntryID = entryId,
                    DurationMinutes = durationMinutes,
                    ActionByUserId = actionByUserId,
                    ActionByUserName = actionByUserName,
                    Remarks = remarks
                });

                return affected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error approving remote access request {EntryID}", entryId);
                return false;
            }
        }

        public async Task<bool> RejectRequestAsync(int entryId, int actionByUserId, string actionByUserName, string? remarks)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    UPDATE RemoteAccessRequests 
                    SET Status = 'Rejected',
                        ActionDT = GETDATE(),
                        ActionByUserID = @ActionByUserId,
                        ActionByUserName = @ActionByUserName,
                        Remarks = @Remarks
                    WHERE EntryID = @EntryID AND Status = 'Pending';";

                var affected = await db.ExecuteAsync(sql, new
                {
                    EntryID = entryId,
                    ActionByUserId = actionByUserId,
                    ActionByUserName = actionByUserName,
                    Remarks = remarks
                });

                return affected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error rejecting remote access request {EntryID}", entryId);
                return false;
            }
        }

        public async Task<int> GetPendingCountAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "SELECT COUNT(*) FROM RemoteAccessRequests WHERE Status = 'Pending'";
                return await db.ExecuteScalarAsync<int>(sql);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting pending remote access requests count");
                return 0;
            }
        }
    }
}

