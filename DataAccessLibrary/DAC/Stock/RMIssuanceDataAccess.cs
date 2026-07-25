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
    public class RMIssuanceDataAccess : IRMIssuanceDataAccess
    {
        private readonly string _connectionString;

        public RMIssuanceDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<IEnumerable<MakerLookupModel>> GetMakersAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = "SELECT VendID, (VendID1 + ' ' + VenderName) AS VenderName FROM Makers ORDER BY VendID1";
            return await db.QueryAsync<MakerLookupModel>(sql);
        }

        public async Task<IEnumerable<EmployeeLookupModel>> GetEmployeesAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = "SELECT EmpID, Name FROM Employees ORDER BY EmpID";
            return await db.QueryAsync<EmployeeLookupModel>(sql);
        }

        public async Task<IEnumerable<RMLookupModel>> GetMaterialsAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = "SELECT RMID, RMID1, RMName, RMSize, SizeUnit, GroupID FROM RM WHERE InActive = 0 ORDER BY RMID";
            return await db.QueryAsync<RMLookupModel>(sql);
        }

        public async Task<IEnumerable<RMLocationModel>> GetMaterialLocationsAsync(int rmid)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = @"
                SELECT Shelf_RefID, 
                       (StoreName + ' [' + ISNULL(RackNo,'') + '] [' + ISNULL(ShelfNo,'') + ']') AS Store, 
                       ShelfQty, SheetsQty, BatchNo, LotNo 
                FROM VMaterialLocationANDPOWiseBalances 
                WHERE RMID = @RMID AND ShelfQty > 0";
            return await db.QueryAsync<RMLocationModel>(sql, new { RMID = rmid });
        }

        public async Task<RMDetailsModel> GetRMDetailsAsync(int rmid)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = "SELECT MinLevel, Rate, Unit, QtyInStock FROM RM WHERE RMID = @RMID";
            return await db.QueryFirstOrDefaultAsync<RMDetailsModel>(sql, new { RMID = rmid });
        }

        public async Task<IEnumerable<RMHistoryRowModel>> GetIssuanceHistoryAsync(bool isMaker, string recipientId, DateTime fromDate, DateTime toDate)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql;
            if (isMaker)
            {
                sql = @"SELECT RMID1, RMName, SUM(Qty) AS Qty 
                        FROM VRawMaterialIssuance 
                        WHERE VendID = @VendID AND DT BETWEEN @FromDate AND @ToDate 
                        GROUP BY RMID1, RMName";
                return await db.QueryAsync<RMHistoryRowModel>(sql, new { VendID = int.Parse(recipientId), FromDate = fromDate, ToDate = toDate });
            }
            else
            {
                sql = @"SELECT RMID1, RMName, SUM(Qty) AS Qty 
                        FROM VRawMaterialIssuance 
                        WHERE IssdToEmpID = @EmpID AND DT BETWEEN @FromDate AND @ToDate 
                        GROUP BY RMID1, RMName";
                return await db.QueryAsync<RMHistoryRowModel>(sql, new { EmpID = recipientId, FromDate = fromDate, ToDate = toDate });
            }
        }

        public async Task<IEnumerable<IssuanceReceiptLookupModel>> GetProductionIssuancesAsync(int vendorId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = @"SELECT TOP 5000 EntryID, RecieptID 
                           FROM VendIssued 
                           WHERE VendID = @VendID 
                           ORDER BY EntryID DESC";
            return await db.QueryAsync<IssuanceReceiptLookupModel>(sql, new { VendID = vendorId });
        }

        public async Task<IEnumerable<RMLookupModel>> GetMaterialsByIssuanceAsync(int entryId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = @"SELECT RMID, RMID1, RMName, RMSize, SizeUnit, GroupID 
                           FROM RM 
                           WHERE InActive = 0 
                             AND RMID IN (SELECT RMID 
                                          FROM ItemsRMComp 
                                          WHERE ItemID = (SELECT ITemID FROM VVendIssued WHERE EntryID = @EntryID) 
                                            AND Functional_Status = 0) 
                           ORDER BY RMID";
            return await db.QueryAsync<RMLookupModel>(sql, new { EntryID = entryId });
        }

        public async Task<string> SaveIssuanceTransactionAsync(RMIssuanceSaveRequest request)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            db.Open();
            using var transaction = db.BeginTransaction();

            try
            {
                // 1. Generate IssNo
                string issNoSql = "SELECT ISNULL(MAX(CAST(RIGHT(IssNo, 3) AS INT)), 0) FROM RawMaterialIssuance WHERE CAST(DT AS DATE) = CAST(@DT AS DATE)";
                int maxNo = await db.ExecuteScalarAsync<int>(issNoSql, new { request.DT }, transaction);
                string newIssNo = $"ISS-{request.DT:yyMMdd}-{(maxNo + 1):000}";

                // 2. Insert Header
                string headerSql = @"
                    INSERT INTO RawMaterialIssuance (IssNo, DT, UserName, ComputerName, VI_RefID, IssdToEmpID, AuthorizedByEmpID, VO_OrderNo, Remarks)
                    VALUES (@IssNo, @DT, @UserName, @ComputerName, @VI_RefID, @IssdToEmpID, @AuthorizedByEmpID, @VO_OrderNo, @Remarks)";
                
                await db.ExecuteAsync(headerSql, new
                {
                    IssNo = newIssNo,
                    request.DT,
                    request.UserName,
                    ComputerName = request.MachineName,
                    request.VI_RefID,
                    request.IssdToEmpID,
                    request.AuthorizedByEmpID,
                    VO_OrderNo = "", // Legacy used it optionally for chkRMPOs
                    request.Remarks
                }, transaction);

                // 3. Insert Details & Update stock
                foreach (var row in request.Rows)
                {
                    string detailSql = @"
                        INSERT INTO RawMaterialIssuanceDetail (IssNo, RMID1, Qty, InHand, StoreName, Price, Shelf_RefID, Sheets, BatchNo, LotNo)
                        VALUES (@IssNo, @RMID1, @Qty, @InHand, @StoreName, @Price, @Shelf_RefID, 0, @BatchNo, @LotNo);
                        SELECT CAST(SCOPE_IDENTITY() as int);";

                    int detailEntryId = await db.ExecuteScalarAsync<int>(detailSql, new
                    {
                        IssNo = newIssNo,
                        row.RMID1,
                        row.Qty,
                        InHand = row.TotalQtyInStock - row.Qty,
                        StoreName = row.Store,
                        Price = row.Rate,
                        row.Shelf_RefID,
                        row.BatchNo,
                        row.LotNo
                    }, transaction);

                    // Update RM QtyInStock
                    string updateRmSql = "UPDATE RM SET QtyInStock = ISNULL(QtyInStock, 0) - @Qty WHERE RMID = @RMID";
                    await db.ExecuteAsync(updateRmSql, new { row.Qty, row.RMID }, transaction);

                    // Update Material Location via SP
                    await db.ExecuteAsync("UpdateMaterialLocation", new
                    {
                        MaterialID = row.RMID1,
                        LocationID = row.Shelf_RefID,
                        Qty = row.Qty,
                        Access = (bool?)null,
                        NewLocationID = (int?)null,
                        UserName = request.UserName,
                        MachineName = request.MachineName,
                        Sheets = 0,
                        PORefNo = "",
                        DetailEntryID = detailEntryId,
                        BatchNo = row.BatchNo,
                        LotNo = row.LotNo
                    }, transaction, commandType: CommandType.StoredProcedure);
                }

                transaction.Commit();
                return newIssNo;
            }
            catch (Exception)
            {
                transaction.Rollback();
                throw;
            }
        }
    }
}
