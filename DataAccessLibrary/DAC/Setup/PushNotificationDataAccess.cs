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
    public class PushNotificationDataAccess : IPushNotificationDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<PushNotificationDataAccess> _logger;

        public PushNotificationDataAccess(IConfiguration configuration, ILogger<PushNotificationDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task UpsertSubscriptionAsync(string userName, int? userId, string endpoint, string p256dh, string auth, string? deviceName)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    IF EXISTS (SELECT 1 FROM UserPushSubscriptions WHERE Endpoint = @Endpoint)
                    BEGIN
                        UPDATE UserPushSubscriptions
                        SET UserID = @UserID,
                            UserName = @UserName,
                            P256dh = @P256dh,
                            Auth = @Auth,
                            DeviceName = ISNULL(@DeviceName, DeviceName),
                            LastUsedAt = GETDATE()
                        WHERE Endpoint = @Endpoint;
                    END
                    ELSE
                    BEGIN
                        INSERT INTO UserPushSubscriptions (UserID, UserName, Endpoint, P256dh, Auth, DeviceName, CreatedAt, LastUsedAt)
                        VALUES (@UserID, @UserName, @Endpoint, @P256dh, @Auth, @DeviceName, GETDATE(), GETDATE());
                    END";

                await db.ExecuteAsync(sql, new
                {
                    UserID = userId,
                    UserName = userName,
                    Endpoint = endpoint,
                    P256dh = p256dh,
                    Auth = auth,
                    DeviceName = deviceName
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error upserting push subscription for user {UserName}", userName);
                throw;
            }
        }

        public async Task<List<UserPushSubscriptionModel>> GetSubscriptionsByUserNameAsync(string userName)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT Id, UserID, UserName, Endpoint, P256dh, Auth, DeviceName, CreatedAt, LastUsedAt
                    FROM UserPushSubscriptions
                    WHERE UserName = @UserName";

                var result = await db.QueryAsync<UserPushSubscriptionModel>(sql, new { UserName = userName });
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching push subscriptions for user {UserName}", userName);
                return new List<UserPushSubscriptionModel>();
            }
        }

        public async Task<List<UserPushSubscriptionModel>> GetSubscriptionsByUserIdAsync(string userId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT Id, UserID, UserName, Endpoint, P256dh, Auth, DeviceName, CreatedAt, LastUsedAt
                    FROM UserPushSubscriptions
                    WHERE UserName = @UserId OR CAST(UserID AS NVARCHAR(50)) = @UserId";

                var result = await db.QueryAsync<UserPushSubscriptionModel>(sql, new { UserId = userId });
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching push subscriptions for user ID {UserId}", userId);
                return new List<UserPushSubscriptionModel>();
            }
        }

        public async Task<List<UserPushSubscriptionModel>> GetAllActiveSubscriptionsAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT Id, UserID, UserName, Endpoint, P256dh, Auth, DeviceName, CreatedAt, LastUsedAt
                    FROM UserPushSubscriptions";

                var result = await db.QueryAsync<UserPushSubscriptionModel>(sql);
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching all push subscriptions");
                return new List<UserPushSubscriptionModel>();
            }
        }

        public async Task DeleteSubscriptionByEndpointAsync(string endpoint)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "DELETE FROM UserPushSubscriptions WHERE Endpoint = @Endpoint";
                await db.ExecuteAsync(sql, new { Endpoint = endpoint });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting expired push subscription endpoint");
            }
        }

        public Task DeactivateSubscriptionAsync(string endpoint) => DeleteSubscriptionByEndpointAsync(endpoint);
    }
}
