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
    public class FavouriteDataAccess : IFavouriteDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<FavouriteDataAccess> _logger;

        public FavouriteDataAccess(IConfiguration configuration, ILogger<FavouriteDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task<List<FavouriteItemModel>> GetUserFavouritesAsync(int userId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT f.UserID, 
                           LTRIM(RTRIM(f.OptionID)) AS OptionID, 
                           ISNULL(m.OptionName, f.OptionID) AS OptionName,
                           ISNULL(m.ModuleName, '') AS ModuleName
                    FROM User_FavouriteLinks f
                    LEFT JOIN MenuOptions m ON LTRIM(RTRIM(f.OptionID)) = LTRIM(RTRIM(m.OptionID))
                    WHERE f.UserID = @UserID
                    ORDER BY OptionName";

                var result = await db.QueryAsync<FavouriteItemModel>(sql, new { UserID = userId });
                return result.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting favourites for UserID {UserID}", userId);
                return new List<FavouriteItemModel>();
            }
        }

        public async Task<bool> AddFavouriteAsync(int userId, string optionId)
        {
            try
            {
                using var db = CreateConnection();
                const string checkSql = "SELECT COUNT(1) FROM User_FavouriteLinks WHERE UserID = @UserID AND OptionID = @OptionID";
                var count = await db.ExecuteScalarAsync<int>(checkSql, new { UserID = userId, OptionID = optionId });
                if (count > 0)
                {
                    return true; // Already exists
                }

                const string insertSql = "INSERT INTO User_FavouriteLinks (UserID, OptionID) VALUES (@UserID, @OptionID)";
                var rows = await db.ExecuteAsync(insertSql, new { UserID = userId, OptionID = optionId });
                return rows > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error adding favourite OptionID {OptionID} for UserID {UserID}", optionId, userId);
                return false;
            }
        }

        public async Task<bool> RemoveFavouriteAsync(int userId, string optionId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "DELETE FROM User_FavouriteLinks WHERE UserID = @UserID AND OptionID = @OptionID";
                var rows = await db.ExecuteAsync(sql, new { UserID = userId, OptionID = optionId });
                return rows > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error removing favourite OptionID {OptionID} for UserID {UserID}", optionId, userId);
                return false;
            }
        }

        public async Task<bool> IsFavouriteAsync(int userId, string optionId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "SELECT COUNT(1) FROM User_FavouriteLinks WHERE UserID = @UserID AND OptionID = @OptionID";
                var count = await db.ExecuteScalarAsync<int>(sql, new { UserID = userId, OptionID = optionId });
                return count > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error checking favourite OptionID {OptionID} for UserID {UserID}", optionId, userId);
                return false;
            }
        }
    }
}
