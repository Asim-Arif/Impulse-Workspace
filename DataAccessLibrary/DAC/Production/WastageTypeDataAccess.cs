using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Production
{
    public class WastageTypeDataAccess : IWastageTypeDataAccess
    {
        private readonly IConfiguration _configuration;

        public WastageTypeDataAccess(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        private string ConnectionString => _configuration.GetConnectionString("DefaultConnection")
            ?? _configuration.GetConnectionString("ImpulseDatabase")
            ?? string.Empty;

        public async Task<List<WastageTypeModel>> GetWastageTypesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            const string sql = @"
                SELECT 
                    EntryID,
                    ISNULL(WastageName, '') AS WastageName,
                    ISNULL(WastageType, 0) AS WastageType,
                    ISNULL(Closed, 0) AS Closed
                FROM WastageTypes
                ORDER BY EntryID";

            var list = (await db.QueryAsync<WastageTypeModel>(sql)).ToList();
            return list;
        }

        public async Task<WastageTypeDto?> GetWastageTypeByIdAsync(int entryId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            const string sql = @"
                SELECT 
                    EntryID,
                    ISNULL(WastageName, '') AS WastageName,
                    ISNULL(WastageType, 0) AS WastageType,
                    ISNULL(Closed, 0) AS Closed
                FROM WastageTypes
                WHERE EntryID = @EntryID";

            return await db.QueryFirstOrDefaultAsync<WastageTypeDto>(sql, new { EntryID = entryId });
        }

        public async Task<int> SaveWastageTypeAsync(WastageTypeDto model)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            const string sql = @"
                INSERT INTO WastageTypes (
                    WastageName, WastageType, Closed
                ) VALUES (
                    @WastageName, @WastageType, @Closed
                );
                SELECT CAST(SCOPE_IDENTITY() AS INT);";

            return await db.ExecuteScalarAsync<int>(sql, new
            {
                WastageName = model.WastageName ?? string.Empty,
                WastageType = model.WastageType,
                Closed = model.Closed ? 1 : 0
            });
        }

        public async Task<bool> UpdateWastageTypeAsync(WastageTypeDto model)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            const string sql = @"
                UPDATE WastageTypes SET 
                    WastageName = @WastageName,
                    WastageType = @WastageType,
                    Closed = @Closed
                WHERE EntryID = @EntryID;";

            int rows = await db.ExecuteAsync(sql, new
            {
                WastageName = model.WastageName ?? string.Empty,
                WastageType = model.WastageType,
                Closed = model.Closed ? 1 : 0,
                EntryID = model.EntryID
            });

            return rows > 0;
        }

        public async Task<(bool CanDelete, string Reason)> CanDeleteWastageTypeAsync(int entryId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            int usageCount = await db.ExecuteScalarAsync<int>(
                "SELECT COUNT(1) FROM VendRcvdDetailWastageDetail WHERE Wastage_RefID = @EntryID",
                new { EntryID = entryId });

            if (usageCount > 0)
                return (false, $"Cannot delete wastage type: It is currently linked to {usageCount} vendor receiving wastage record(s).");

            return (true, string.Empty);
        }

        public async Task<bool> DeleteWastageTypeAsync(int entryId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            int rows = await db.ExecuteAsync(
                "DELETE FROM WastageTypes WHERE EntryID = @EntryID",
                new { EntryID = entryId });

            return rows > 0;
        }
    }
}
