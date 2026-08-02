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
            ?? throw new InvalidOperationException("DefaultConnection configuration string is missing.");

        public async Task<List<MakerLookupModel>> GetActiveMakersAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT VendID, VendID1, VenderName 
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

        public async Task<List<UnassignedItemModel>> GetUnassignedItemsAsync(long vendId, int processId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT ItemID, ItemName, ItemSize, SizeUnit, FinQuality 
                           FROM VItems 
                           WHERE ItemID IN (SELECT ItemID FROM ItemProcesses WHERE ProcessID = @ProcessId)
                             AND ItemID NOT IN (SELECT ItemID FROM VendAssItems WHERE VendID = @VendId AND ProcessID = @ProcessId)
                           ORDER BY ItemName";

            return (await db.QueryAsync<UnassignedItemModel>(sql, new { VendId = vendId, ProcessId = processId })).ToList();
        }

        public async Task<List<AssignedMakerItemModel>> GetAssignedItemsAsync(long vendId, int processId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT EntryID, VendID, ProcessID, ItemID, ItemName, ISNULL(Description, '') AS Description, 
                                  ISNULL(Rate, 0) AS Rate, ISNULL(AssignedUnit, '') AS AssignedUnit, 
                                  ISNULL(Remarks, '') AS Remarks, ISNULL(MakerDescription, '') AS MakerDescription
                           FROM VMakerAssItems
                           WHERE VendID = @VendId AND ProcessID = @ProcessId
                           ORDER BY ItemName";

            return (await db.QueryAsync<AssignedMakerItemModel>(sql, new { VendId = vendId, ProcessId = processId })).ToList();
        }

        public async Task AssignItemAsync(long vendId, int processId, string itemId, decimal rate)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            
            // Execute SP_VendAssItem stored procedure matching legacy form
            var parameters = new DynamicParameters();
            parameters.Add("@VendID", vendId);
            parameters.Add("@ItemID", itemId);
            parameters.Add("@ProcessID", processId);
            parameters.Add("@Rate", rate);

            await db.ExecuteAsync("SP_VendAssItem", parameters, commandType: CommandType.StoredProcedure);
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
            string sql = @"SELECT EntryID, VAI_RefID, Rate, ISNULL(UserName, '') AS UserName, DTEntry, ISNULL(Remarks, '') AS Remarks 
                           FROM VendAssItems_Revisions 
                           WHERE VAI_RefID = @EntryId 
                           ORDER BY EntryID DESC";

            return (await db.QueryAsync<ItemRevisionHistoryModel>(sql, new { EntryId = entryId })).ToList();
        }
    }
}
