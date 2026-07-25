using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Models.ViewModels.Stock;
using DataAccessLibrary.Interface.Stock;

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Data.SqlClient;
using System.Linq;

namespace DataAccessLibrary.DAC.Stock
{
    public class ChangeBatchNoDataAccess : IChangeBatchNoDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<ChangeBatchNoDataAccess> _logger;

        public ChangeBatchNoDataAccess(IConfiguration configuration, ILogger<ChangeBatchNoDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                             ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        public async Task<IEnumerable<object>> GetMaterialsAsync()
        {
            using IDbConnection connection = new SqlConnection(_connectionString);
            var sql = "SELECT RMID1, RMName, '{' + RMID1 + '} ' + RMName AS DisplayName FROM RM ORDER BY RMID1";
            return await connection.QueryAsync<dynamic>(sql);
        }

        public async Task<IEnumerable<ChangeBatchNoViewModel>> GetMaterialBatchesAsync(string materialId)
        {
            var sql = @"
                SELECT EntryID, RcvID, AccTitle AS VendorName, QtyPlaced, StoreName, RackNo, ShelfNo, 
                       BatchNo, LotNo, Mill_Certificate_No AS MillCertificateNo 
                FROM VMaterialLocationWiseStatus 
                WHERE MaterialID = @MaterialId 
                ORDER BY RcvDate, EntryID";

            using IDbConnection connection = new SqlConnection(_connectionString);
            var data = await connection.QueryAsync<ChangeBatchNoViewModel>(sql, new { MaterialId = materialId });
            foreach (var item in data)
            {
                item.OriginalBatchNo = item.BatchNo;
                item.OriginalLotNo = item.LotNo;
                item.OriginalMillCertificateNo = item.MillCertificateNo;
            }
            return data;
        }

        public async Task SaveBatchModificationsAsync(IEnumerable<ChangeBatchNoViewModel> modifiedItems)
        {
            using IDbConnection connection = new SqlConnection(_connectionString);
            connection.Open();

            using var transaction = connection.BeginTransaction();
            try
            {
                foreach (var item in modifiedItems)
                {
                    // Update Lots_List
                    if (!string.IsNullOrEmpty(item.OriginalBatchNo) && item.BatchNo != item.OriginalBatchNo)
                    {
                        var updateBatchSql = "UPDATE Lots_List SET Batch_No = @BatchNo WHERE Batch_No = @OriginalBatchNo";
                        await connection.ExecuteAsync(updateBatchSql, new { BatchNo = item.BatchNo, OriginalBatchNo = item.OriginalBatchNo }, transaction);
                    }

                    if (!string.IsNullOrEmpty(item.OriginalMillCertificateNo) && item.MillCertificateNo != item.OriginalMillCertificateNo)
                    {
                        var updateMillCertSql = "UPDATE Lots_List SET Mill_Certificate_No = @MillCert WHERE Mill_Certificate_No = @OriginalMillCert";
                        await connection.ExecuteAsync(updateMillCertSql, new { MillCert = item.MillCertificateNo, OriginalMillCert = item.OriginalMillCertificateNo }, transaction);
                    }

                    // Update MaterialLocationwiseStatus
                    var updateStatusSql = @"
                        UPDATE MaterialLocationwiseStatus 
                        SET BatchNo = @BatchNo, 
                            LotNo = @LotNo, 
                            Mill_Certificate_No = @MillCert 
                        WHERE EntryID = @EntryID";
                        
                    await connection.ExecuteAsync(updateStatusSql, new 
                    { 
                        BatchNo = item.BatchNo ?? "", 
                        LotNo = item.LotNo ?? "", 
                        MillCert = item.MillCertificateNo ?? "", 
                        EntryID = item.EntryID 
                    }, transaction);
                }

                transaction.Commit();
            }
            catch
            {
                transaction.Rollback();
                throw;
            }
        }
    }
}
