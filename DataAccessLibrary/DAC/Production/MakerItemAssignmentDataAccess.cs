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
    public class MakerItemAssignmentDataAccess : IMakerItemAssignmentDataAccess
    {
        private readonly IConfiguration _config;

        public MakerItemAssignmentDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? _config.GetConnectionString("ImpulseConnection")
            ?? _config.GetConnectionString("ImpulseDatabase")
            ?? throw new InvalidOperationException("DefaultConnection configuration string is missing.");

        private static bool? _hasCustCodeColumn;

        private async Task<bool> HasCustCodeColumnAsync(IDbConnection db)
        {
            if (_hasCustCodeColumn.HasValue) return _hasCustCodeColumn.Value;
            try
            {
                int count = await db.ExecuteScalarAsync<int>(
                    "SELECT COUNT(1) FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'VendAssItems' AND COLUMN_NAME = 'CustCode'");
                _hasCustCodeColumn = count > 0;
            }
            catch
            {
                _hasCustCodeColumn = false;
            }
            return _hasCustCodeColumn.Value;
        }

        public async Task<bool> HasCustomerWiseRatesSupportAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            bool colExists = await HasCustCodeColumnAsync(db);
            bool configEnabled = bool.TryParse(_config["AppSettings:EnableCustomerWiseMakerRates"], out bool enabled) && enabled;
            return colExists && configEnabled;
        }

        public async Task<List<CustomerLookupModel>> GetCustomersAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT DISTINCT CustCode, ISNULL(Name, CustCode) AS Name 
                           FROM ForeignCustomers 
                           WHERE CustCode IS NOT NULL AND CustCode <> ''
                           ORDER BY CustCode";

            return (await db.QueryAsync<CustomerLookupModel>(sql)).ToList();
        }

        public async Task<List<MakerLookupModel>> GetActiveMakersAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT CAST(VendID AS bigint) AS VendID, VendID1, VenderName 
                           FROM VMakers 
                           WHERE ISNULL(Active, 1) = 1 
                           ORDER BY VenderName";

            return (await db.QueryAsync<MakerLookupModel>(sql)).ToList();
        }

        public async Task<List<ProcessLookupModel>> GetMakerAssignedProcessesAsync(long vendId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT ProcessID, Description, SNO 
                           FROM VMakerAssignedProcesses 
                           WHERE VendID = @VendId 
                           ORDER BY SNO";

            return (await db.QueryAsync<ProcessLookupModel>(sql, new { VendId = vendId })).ToList();
        }

        public async Task<List<UnassignedItemModel>> GetUnassignedItemsAsync(long vendId, int processId, string? custCode = null)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            bool isCustWise = await HasCustomerWiseRatesSupportAsync();
            string custFilter = isCustWise && !string.IsNullOrWhiteSpace(custCode)
                ? " AND CustCode = @CustCode"
                : "";

            string sql = $@"SELECT ItemID, ItemName, ItemSize, SizeUnit, FinQuality 
                           FROM VItems 
                           WHERE ItemID IN (SELECT ItemID FROM ItemProcesses WHERE ProcessID = @ProcessId)
                             AND ItemID NOT IN (SELECT ItemID FROM VendAssItems WHERE VendID = @VendId AND ProcessID = @ProcessId{custFilter})
                           ORDER BY ItemName";

            return (await db.QueryAsync<UnassignedItemModel>(sql, new { VendId = vendId, ProcessId = processId, CustCode = custCode })).ToList();
        }

        public async Task<List<AssignedMakerItemModel>> GetAssignedItemsAsync(long vendId, int processId, string? custCode = null)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            bool hasCust = await HasCustCodeColumnAsync(db);
            bool isCustWise = await HasCustomerWiseRatesSupportAsync();
            string custSelect = hasCust ? "v.CustCode," : "CAST(NULL AS varchar(50)) AS CustCode,";

            string sql;
            if (isCustWise && !string.IsNullOrWhiteSpace(custCode))
            {
                sql = $@"SELECT CAST(v.EntryID AS bigint) AS EntryID, 
                                CAST(v.VendID AS bigint) AS VendID, 
                                v.ProcessID, 
                                v.ItemID, 
                                v.ItemName, 
                                ISNULL(v.Description, '') AS Description, 
                                CAST(ISNULL(v.Rate, 0) AS decimal(18, 4)) AS Rate, 
                                ISNULL(v.AssignedUnit, '') AS AssignedUnit, 
                                ISNULL(v.Remarks, '') AS Remarks, 
                                ISNULL(i.MakerDescription, '') AS MakerDescription,
                                {custSelect}
                                CAST(v.EntryID AS bigint) AS SortEntryID
                         FROM VMakerAssItems v
                         LEFT OUTER JOIN Items i ON v.ItemID = i.ItemID
                         WHERE v.VendID = @VendId AND v.ProcessID = @ProcessId AND v.CustCode = @CustCode
                         ORDER BY v.ItemName";
            }
            else
            {
                // When CustomerWise rates is disabled or no customer filter is applied,
                // deduplicate items so each ItemID appears only once, prioritizing the active/highest rate and latest entry.
                sql = $@"WITH RankedAssignedItems AS (
                            SELECT CAST(v.EntryID AS bigint) AS EntryID, 
                                   CAST(v.VendID AS bigint) AS VendID, 
                                   v.ProcessID, 
                                   v.ItemID, 
                                   v.ItemName, 
                                   ISNULL(v.Description, '') AS Description, 
                                   CAST(ISNULL(v.Rate, 0) AS decimal(18, 4)) AS Rate, 
                                   ISNULL(v.AssignedUnit, '') AS AssignedUnit, 
                                   ISNULL(v.Remarks, '') AS Remarks, 
                                   ISNULL(i.MakerDescription, '') AS MakerDescription,
                                   {custSelect}
                                   CAST(v.EntryID AS bigint) AS SortEntryID,
                                   ROW_NUMBER() OVER (PARTITION BY v.ItemID ORDER BY v.Rate DESC, v.EntryID DESC) as rn
                            FROM VMakerAssItems v
                            LEFT OUTER JOIN Items i ON v.ItemID = i.ItemID
                            WHERE v.VendID = @VendId AND v.ProcessID = @ProcessId
                         )
                         SELECT EntryID, VendID, ProcessID, ItemID, ItemName, Description, Rate, AssignedUnit, Remarks, MakerDescription, CustCode, SortEntryID
                         FROM RankedAssignedItems
                         WHERE rn = 1
                         ORDER BY ItemName";
            }

            return (await db.QueryAsync<AssignedMakerItemModel>(sql, new { VendId = vendId, ProcessId = processId, CustCode = custCode })).ToList();
        }

        public async Task AssignItemAsync(long vendId, int processId, string itemId, decimal rate, string? custCode = null)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            bool isCustWise = await HasCustomerWiseRatesSupportAsync();
            
            // Execute SP_VendAssItem stored procedure matching legacy form
            var parameters = new DynamicParameters();
            parameters.Add("@VendID", vendId);
            parameters.Add("@ItemID", itemId);
            parameters.Add("@ProcessID", processId);
            parameters.Add("@Rate", rate);
            if (isCustWise && !string.IsNullOrWhiteSpace(custCode))
            {
                parameters.Add("@CustCode", custCode);
            }

            await db.ExecuteAsync("SP_VendAssItem", parameters, commandType: CommandType.StoredProcedure);
        }

        public async Task<int> AssignAllUnassignedItemsAsync(long vendId, int processId, decimal rate, string? custCode = null)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            bool isCustWise = await HasCustomerWiseRatesSupportAsync();
            string effectiveCustCode = isCustWise && !string.IsNullOrWhiteSpace(custCode) ? custCode.Trim() : "Stock";

            string countSql = @"SELECT COUNT(1)
                                FROM Items
                                WHERE ItemID IN (SELECT ItemID FROM ItemProcesses WHERE ProcessID = @ProcessID)
                                  AND ItemID NOT IN (
                                      SELECT ItemID
                                      FROM VendAssItems
                                      WHERE VendID = @VendID
                                        AND ProcessID = @ProcessID
                                        AND (CustCode = @EffectiveCustCode OR (@EffectiveCustCode = 'Stock' AND (CustCode IS NULL OR CustCode = '' OR CustCode = 'Stock')))
                                  )";

            int pendingCount = await db.ExecuteScalarAsync<int>(countSql, new
            {
                VendID = vendId,
                ProcessID = processId,
                EffectiveCustCode = effectiveCustCode
            });

            if (pendingCount <= 0)
            {
                return 0;
            }

            var parameters = new DynamicParameters();
            parameters.Add("@VendID", vendId);
            parameters.Add("@ItemID", "0");
            parameters.Add("@ProcessID", processId);
            parameters.Add("@Rate", rate);
            if (isCustWise && !string.IsNullOrWhiteSpace(custCode))
            {
                parameters.Add("@CustCode", custCode.Trim());
            }

            await db.ExecuteAsync("SP_VendAssItem", parameters, commandType: CommandType.StoredProcedure);
            return pendingCount;
        }

        public async Task UpdateRateForAllCustomersAsync(long vendId, int processId, string itemId, decimal newRate, decimal oldRate, string newRemarks, string userName)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                string updateSql = @"UPDATE VendAssItems 
                                     SET Rate = @NewRate, Remarks = @Remarks 
                                     WHERE VendID = @VendId AND ProcessID = @ProcessId AND ItemID = @ItemId";

                await db.ExecuteAsync(updateSql, new { VendId = vendId, ProcessId = processId, ItemId = itemId, NewRate = newRate, Remarks = newRemarks }, trans);

                // Insert into revisions log for all matching items if rate changed
                if (newRate != oldRate)
                {
                    string getEntriesSql = @"SELECT EntryID FROM VendAssItems WHERE VendID = @VendId AND ProcessID = @ProcessId AND ItemID = @ItemId";
                    var entryIds = (await db.QueryAsync<long>(getEntriesSql, new { VendId = vendId, ProcessId = processId, ItemId = itemId }, trans)).ToList();

                    foreach (var eid in entryIds)
                    {
                        string revSql = @"INSERT INTO VendAssItems_Revisions (VAI_RefID, Rate, UserName, Remarks) 
                                          VALUES (@EntryId, @OldRate, @UserName, @Remarks)";
                        await db.ExecuteAsync(revSql, new { EntryId = eid, OldRate = oldRate, UserName = userName, Remarks = newRemarks }, trans);
                    }
                }

                trans.Commit();
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }

        public async Task UpdateAssignedItemRateAndRemarksAsync(long entryId, decimal newRate, decimal oldRate, string newRemarks, string userName)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                // Update item assignment
                string updateSql = @"UPDATE VendAssItems 
                                     SET Rate = @NewRate, Remarks = @Remarks 
                                     WHERE EntryID = @EntryId";

                await db.ExecuteAsync(updateSql, new { EntryId = entryId, NewRate = newRate, Remarks = newRemarks }, trans);

                // Insert into revisions log if rate changed
                if (newRate != oldRate)
                {
                    string revSql = @"INSERT INTO VendAssItems_Revisions (VAI_RefID, Rate, UserName, Remarks) 
                                      VALUES (@EntryId, @OldRate, @UserName, @Remarks)";

                    await db.ExecuteAsync(revSql, new { EntryId = entryId, OldRate = oldRate, UserName = userName, Remarks = newRemarks }, trans);
                }

                trans.Commit();
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }

        public async Task<bool> IsItemReceivableAsync(long vendId, int processId, string itemId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT COUNT(1) 
                           FROM VVendIssdDetail 
                           WHERE VendID = @VendId AND ProcessID = @ProcessId AND ItemCode = @ItemId AND ISNULL(IssQty, 0) > ISNULL(RcvdQty, 0)";

            int count = await db.ExecuteScalarAsync<int>(sql, new { VendId = vendId, ProcessId = processId, ItemId = itemId });
            return count > 0;
        }

        public async Task UnassignItemAsync(long entryId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"DELETE FROM VendAssItems WHERE EntryID = @EntryId";
            await db.ExecuteAsync(sql, new { EntryId = entryId });
        }

        public async Task<List<ItemRevisionHistoryModel>> GetItemRevisionHistoryAsync(long entryId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT CAST(EntryID AS bigint) AS EntryID, 
                                  CAST(VAI_RefID AS bigint) AS VAI_RefID, 
                                  CAST(ISNULL(Rate, 0) AS decimal(18, 4)) AS Rate, 
                                  ISNULL(UserName, '') AS UserName, 
                                  DTEntry, 
                                  ISNULL(Remarks, '') AS Remarks 
                           FROM VendAssItems_Revisions 
                           WHERE VAI_RefID = @EntryId 
                           ORDER BY EntryID DESC";

            return (await db.QueryAsync<ItemRevisionHistoryModel>(sql, new { EntryId = entryId })).ToList();
        }
    }
}
