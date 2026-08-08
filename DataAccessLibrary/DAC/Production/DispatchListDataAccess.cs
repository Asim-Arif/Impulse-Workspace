using Dapper;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Production
{
    public class DispatchListDataAccess : IDispatchListDataAccess
    {
        private readonly IConfiguration _config;

        public DispatchListDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? _config.GetConnectionString("ImpulseConnection")
            ?? string.Empty;

        public async Task<List<LookupItemString>> GetCustomersAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT '0' AS Id, '<All Customers>' AS Name
                UNION ALL
                SELECT DISTINCT CustCode AS Id, CustCode AS Name
                FROM ForeignCustomers
                WHERE CustCode <> 'Stock'
                ORDER BY Name";
            return (await db.QueryAsync<LookupItemString>(sql)).ToList();
        }

        public async Task<List<DispatchListSummaryItem>> GetDispatchListsAsync(DispatchListSearchFilter filter)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            string sql = @"
                SELECT DispatchListNo, DT, TotalCartons, EntryID, NewFormat,
                       Finalyzed, FinalyzedDT, FinalyzedUserName, CustCode, TotalQty,
                       UserName, EntryDT, Country
                FROM VDispatchList
                WHERE DT >= @DtFrom AND DT <= @DtTo";

            var p = new DynamicParameters();
            p.Add("@DtFrom", filter.DtFrom.Date);
            p.Add("@DtTo", filter.DtTo.Date.AddDays(1).AddTicks(-1));

            if (!string.IsNullOrWhiteSpace(filter.CustCode) && filter.CustCode != "0")
            {
                sql += " AND CustCode = @CustCode";
                p.Add("@CustCode", filter.CustCode.Trim());
            }

            if (filter.OnlyNotFinalized)
            {
                sql += " AND Finalyzed = 0";
            }

            sql += " ORDER BY DT DESC, EntryID DESC";

            return (await db.QueryAsync<DispatchListSummaryItem>(sql, p)).ToList();
        }

        public async Task<FinalizeDispatchResult> FinalizeDispatchListAsync(long entryId, string userName, string machineName)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            // Check if already finalized
            bool isAlreadyFinal = await db.ExecuteScalarAsync<bool>(
                "SELECT ISNULL(Finalyzed, 0) FROM DispatchList WHERE EntryID = @EntryID",
                new { EntryID = entryId });

            if (isAlreadyFinal)
            {
                return new FinalizeDispatchResult { Success = false, Message = "This dispatch list is already finalized." };
            }

            // Check if unassigned cartons exist
            const string countSql = @"
                SELECT COUNT(*)
                FROM DispatchList
                INNER JOIN DispatchListDetails_Adv ON DispatchList.EntryID = DispatchListDetails_Adv.RefID
                INNER JOIN DispatchListDetail_Inners ON DispatchListDetails_Adv.EntryID = DispatchListDetail_Inners.RefID
                WHERE DispatchList.EntryID = @EntryID AND ISNULL(CartonNo, 0) = 0";

            int unassignedCount = await db.ExecuteScalarAsync<int>(countSql, new { EntryID = entryId });
            if (unassignedCount > 0)
            {
                return new FinalizeDispatchResult
                {
                    Success = false,
                    Message = "Can't Finalize, Carton No. not Assigned."
                };
            }

            const string updateSql = @"
                UPDATE DispatchList
                SET Finalyzed = 1,
                    FinalyzedDT = GETDATE(),
                    FinalyzedUserName = @UserName,
                    FinalyzedMachineName = @MachineName
                WHERE EntryID = @EntryID";

            await db.ExecuteAsync(updateSql, new
            {
                UserName = userName,
                MachineName = machineName,
                EntryID = entryId
            });

            return new FinalizeDispatchResult
            {
                Success = true,
                Message = "Dispatch List finalized successfully."
            };
        }
    }
}
