using Dapper;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Stock
{
    public class RMMovementDataAccess : IRMMovementDataAccess
    {
        private readonly string _connectionString;

        public RMMovementDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<List<LocationwiseStatusViewModel>> GetLocationsForMaterialAsync(string materialId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT m.EntryID, m.Shelf_RefID, s.StoreName, s.ShelfNo, 
                           ROUND(m.QtyPlaced - ISNULL(m.QtyIssued, 0), 2) AS AvailableQty, 
                           m.Rcvd_RefID, m.BatchNo, m.LotNo
                    FROM MaterialLocationwiseStatus m
                    INNER JOIN VendRcvdDetailPO v ON m.Rcvd_RefID = v.EntryID
                    INNER JOIN VStoreShelfs s ON m.Shelf_RefID = s.EntryID
                    WHERE v.MaterialID = @MaterialId 
                      AND ROUND(m.QtyPlaced, 2) > ROUND(ISNULL(m.QtyIssued, 0), 2)";
                      
                var results = await db.QueryAsync<LocationwiseStatusViewModel>(sql, new { MaterialId = materialId });
                return results.AsList();
            }
        }

        public async Task<double> GetShelfQuantityAsync(int shelfId, string materialId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT ROUND(SUM(QtyPlaced - ISNULL(QtyIssued, 0)), 2) AS ShelfQty
                    FROM VMaterialLocationWiseStatus
                    WHERE Shelf_RefID = @ShelfId AND MaterialID = @MaterialId AND (QtyPlaced - ISNULL(QtyIssued, 0) > 0)";
                    
                var result = await db.QuerySingleOrDefaultAsync<double?>(sql, new { ShelfId = shelfId, MaterialId = materialId });
                return result ?? 0;
            }
        }

        public async Task SaveMovementAsync(RMMovementSaveRequest request, string userName, string machineName)
        {
            if (request.QtyMoved <= 0)
                throw new InvalidOperationException("Quantity to move must be greater than zero.");

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                db.Open();
                using (var transaction = db.BeginTransaction())
                {
                    try
                    {
                        // 0. Lock source shelf row and verify available balance
                        string checkSql = @"
                            SELECT ROUND(CAST(ISNULL(QtyPlaced, 0) AS FLOAT) - CAST(ISNULL(QtyIssued, 0) AS FLOAT), 4)
                            FROM MaterialLocationwiseStatus WITH (UPDLOCK, HOLDLOCK)
                            WHERE EntryID = @EntryID";

                        double? availableQty = await db.QuerySingleOrDefaultAsync<double?>(
                            checkSql, new { EntryID = request.From_MLS_RefID }, transaction);

                        if (!availableQty.HasValue || request.QtyMoved > availableQty.Value + 0.001)
                        {
                            throw new InvalidOperationException(
                                $"Insufficient stock on source shelf. Requested: {request.QtyMoved:0.####}, Available: {Math.Max(0, availableQty ?? 0):0.####}.");
                        }

                        // 1. Insert new MaterialLocationwiseStatus for the TO location
                        string insertSql = @"
                            INSERT INTO MaterialLocationwiseStatus
                            (Rcvd_RefID, Shelf_RefID, QtyPlaced, QtyIssued, UserName, MachineName, BatchNo, LotNo) 
                            OUTPUT INSERTED.EntryID
                            VALUES
                            (@Rcvd_RefID, @Shelf_RefID, @QtyPlaced, 0, @UserName, @MachineName, @BatchNo, @LotNo)";

                        int newMlsRefId = await db.QuerySingleAsync<int>(insertSql, new
                        {
                            request.Rcvd_RefID,
                            Shelf_RefID = request.To_Shelf_RefID,
                            QtyPlaced = request.QtyMoved,
                            UserName = userName ?? "",
                            MachineName = machineName ?? "",
                            BatchNo = request.BatchNo ?? "",
                            LotNo = request.LotNo ?? ""
                        }, transaction);

                        // 3. Update the QtyIssued on the FROM location
                        string updateSql = @"
                            UPDATE MaterialLocationwiseStatus 
                            SET QtyIssued = ROUND(ISNULL(QtyIssued, 0) + @QtyMoved, 4)
                            WHERE EntryID = @EntryID";

                        await db.ExecuteAsync(updateSql, new
                        {
                            request.QtyMoved,
                            EntryID = request.From_MLS_RefID
                        }, transaction);

                        // 4. Insert into MaterialMovements
                        string insertMovementSql = @"
                            INSERT INTO MaterialMovements
                            (MLS_From_RefID, MLS_To_RefID, UserName, MachineName, QtyMoved) 
                            VALUES
                            (@From_MLS_RefID, @To_MLS_RefID, @UserName, @MachineName, @QtyMoved)";

                        await db.ExecuteAsync(insertMovementSql, new
                        {
                            request.From_MLS_RefID,
                            To_MLS_RefID = newMlsRefId,
                            UserName = userName ?? "",
                            MachineName = machineName ?? "",
                            request.QtyMoved
                        }, transaction);

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
    }
}
