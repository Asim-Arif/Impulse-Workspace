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
    public class ProductionDeletionDataAccess : IProductionDeletionDataAccess
    {
        private readonly IConfiguration _config;

        public ProductionDeletionDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection string is not configured.");

        public async Task<int> CreateRequestAsync(ProductionDeletionRequestModel model)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                INSERT INTO dbo.ProductionDeletionRequests
                    (RequestType, EntityRefID, LotNo, OrderNo, ItemCode, ItemName, 
                     ProcessID, ProcessName, MakerID, MakerName, Qty, RequestedBy, 
                     RequestedDT, Reason, MachineName, Status, TaskId)
                OUTPUT INSERTED.Id
                VALUES
                    (@RequestType, @EntityRefID, @LotNo, @OrderNo, @ItemCode, @ItemName, 
                     @ProcessID, @ProcessName, @MakerID, @MakerName, @Qty, @RequestedBy, 
                     GETDATE(), @Reason, @MachineName, 'Pending', @TaskId);";

            return await db.ExecuteScalarAsync<int>(sql, model);
        }

        public async Task<List<ProductionDeletionRequestModel>> GetPendingRequestsAsync(string? requestType = null)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"
                SELECT Id, RequestType, EntityRefID, LotNo, OrderNo, ItemCode, ItemName, 
                       ProcessID, ProcessName, MakerID, MakerName, Qty, RequestedBy, 
                       RequestedDT, Reason, MachineName, Status, ReviewedBy, ReviewedDT, 
                       DirectorRemarks, TaskId
                FROM dbo.ProductionDeletionRequests
                WHERE Status = 'Pending' " +
                (!string.IsNullOrWhiteSpace(requestType) ? " AND RequestType = @RequestType " : "") +
                " ORDER BY RequestedDT DESC";

            return (await db.QueryAsync<ProductionDeletionRequestModel>(sql, new { RequestType = requestType })).ToList();
        }

        public async Task<ProductionDeletionRequestModel?> GetRequestByIdAsync(int id)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT Id, RequestType, EntityRefID, LotNo, OrderNo, ItemCode, ItemName, 
                       ProcessID, ProcessName, MakerID, MakerName, Qty, RequestedBy, 
                       RequestedDT, Reason, MachineName, Status, ReviewedBy, ReviewedDT, 
                       DirectorRemarks, TaskId
                FROM dbo.ProductionDeletionRequests
                WHERE Id = @Id";

            return await db.QueryFirstOrDefaultAsync<ProductionDeletionRequestModel>(sql, new { Id = id });
        }

        public async Task<ProductionDeletionRequestModel?> GetPendingRequestByEntityIdAsync(string requestType, long entityRefId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT TOP 1 Id, RequestType, EntityRefID, LotNo, OrderNo, ItemCode, ItemName, 
                       ProcessID, ProcessName, MakerID, MakerName, Qty, RequestedBy, 
                       RequestedDT, Reason, MachineName, Status, ReviewedBy, ReviewedDT, 
                       DirectorRemarks, TaskId
                FROM dbo.ProductionDeletionRequests
                WHERE RequestType = @RequestType 
                  AND EntityRefID = @EntityRefId 
                  AND Status = 'Pending'";

            return await db.QueryFirstOrDefaultAsync<ProductionDeletionRequestModel>(sql, new { RequestType = requestType, EntityRefId = entityRefId });
        }

        public async Task<ProductionDeletionRequestModel?> GetPendingRequestByLotNoAsync(string lotNo)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT TOP 1 Id, RequestType, EntityRefID, LotNo, OrderNo, ItemCode, ItemName, 
                       ProcessID, ProcessName, MakerID, MakerName, Qty, RequestedBy, 
                       RequestedDT, Reason, MachineName, Status, ReviewedBy, ReviewedDT, 
                       DirectorRemarks, TaskId
                FROM dbo.ProductionDeletionRequests
                WHERE LotNo = @LotNo 
                  AND Status = 'Pending'";

            return await db.QueryFirstOrDefaultAsync<ProductionDeletionRequestModel>(sql, new { LotNo = lotNo });
        }

        public async Task<bool> IsLotLockedAsync(string lotNo, long vrdEntryId = 0)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT dbo.fn_IsLotPendingDeletion(@LotNo, @VRD_EntryID)";

            return await db.ExecuteScalarAsync<bool>(sql, new { LotNo = lotNo, VRD_EntryID = vrdEntryId });
        }

        public async Task<List<long>> GetActivePendingVRDEntryIdsAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT EntityRefID
                FROM dbo.ProductionDeletionRequests
                WHERE Status = 'Pending' AND RequestType = 'LotReceiving'";

            return (await db.QueryAsync<long>(sql)).ToList();
        }

        public async Task<List<long>> GetActivePendingIssuanceEntryIdsAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT EntityRefID
                FROM dbo.ProductionDeletionRequests
                WHERE Status = 'Pending' AND RequestType IN ('LotIssuance', 'MasterPOIssuance')";

            return (await db.QueryAsync<long>(sql)).ToList();
        }

        public async Task<List<string>> GetActivePendingSkipProcessLotNosAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT DISTINCT LotNo
                FROM dbo.ProductionDeletionRequests WITH (NOLOCK)
                WHERE Status = 'Pending' AND RequestType = 'SkipProcess' AND LotNo IS NOT NULL AND LotNo <> ''";

            return (await db.QueryAsync<string>(sql)).ToList();
        }

        public async Task<List<string>> GetActivePendingDeletionLotNosAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT DISTINCT LotNo
                FROM dbo.ProductionDeletionRequests WITH (NOLOCK)
                WHERE Status = 'Pending' AND RequestType IN ('LotIssuance', 'LotReceiving', 'MasterPOIssuance') AND LotNo IS NOT NULL AND LotNo <> ''";

            return (await db.QueryAsync<string>(sql)).ToList();
        }

        public async Task<bool> IsIssuanceLockedAsync(long issuanceEntryId, string? lotNo = null)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT CASE WHEN EXISTS (
                    SELECT 1 FROM dbo.ProductionDeletionRequests WITH (NOLOCK)
                    WHERE Status = 'Pending'
                      AND RequestType IN ('LotIssuance', 'MasterPOIssuance')
                      AND (
                          (@IssuanceEntryID > 0 AND EntityRefID = @IssuanceEntryID)
                          OR (@LotNo IS NOT NULL AND @LotNo <> '' AND LotNo = @LotNo)
                      )
                ) THEN 1 ELSE 0 END";

            return await db.ExecuteScalarAsync<bool>(sql, new { IssuanceEntryID = issuanceEntryId, LotNo = lotNo ?? "" });
        }

        public async Task<bool> ApproveRequestAsync(int id, string directorUserName, string? remarks = null)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                UPDATE dbo.ProductionDeletionRequests
                SET Status = 'Approved',
                    ReviewedBy = @ReviewedBy,
                    ReviewedDT = GETDATE(),
                    DirectorRemarks = @DirectorRemarks
                WHERE Id = @Id AND Status = 'Pending'";

            var rows = await db.ExecuteAsync(sql, new
            {
                Id = id,
                ReviewedBy = directorUserName,
                DirectorRemarks = remarks
            });

            return rows > 0;
        }

        public async Task<bool> RejectRequestAsync(int id, string directorUserName, string remarks)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                UPDATE dbo.ProductionDeletionRequests
                SET Status = 'Rejected',
                    ReviewedBy = @ReviewedBy,
                    ReviewedDT = GETDATE(),
                    DirectorRemarks = @DirectorRemarks
                WHERE Id = @Id AND Status = 'Pending'";

            var rows = await db.ExecuteAsync(sql, new
            {
                Id = id,
                ReviewedBy = directorUserName,
                DirectorRemarks = remarks
            });

            return rows > 0;
        }

        public async Task<bool> UpdateTaskIdAsync(int requestId, int taskId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                UPDATE dbo.ProductionDeletionRequests
                SET TaskId = @TaskId
                WHERE Id = @Id";

            var rows = await db.ExecuteAsync(sql, new { Id = requestId, TaskId = taskId });
            return rows > 0;
        }
    }
}
