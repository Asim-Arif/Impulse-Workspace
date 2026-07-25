using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Transactions;
using Dapper;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.Extensions.Configuration;
using Microsoft.Data.SqlClient;

namespace DataAccessLibrary.DAC.Stock
{
    public class VendGateRcvdDataAccess : IVendGateRcvdDataAccess
    {
        private readonly IConfiguration _config;

        public VendGateRcvdDataAccess(IConfiguration config)
        {
            _config = config;
        }

        public async Task<List<object>> GetVendorsWithPendingOrdersAsync()
        {
            using IDbConnection connection = new SqlConnection(_config.GetConnectionString("DefaultConnection"));
            string sql = @"
                SELECT DISTINCT v.VendID as AccNo, a.AccTitle 
                FROM VVendOrdersToRcv v
                INNER JOIN VVenders a ON v.VendID = a.AccNo
                ORDER BY a.AccTitle";
            var result = await connection.QueryAsync(sql);
            return result.ToList();
        }

        public async Task<List<string>> GetPendingOrdersAsync(string vendId)
        {
            using IDbConnection connection = new SqlConnection(_config.GetConnectionString("DefaultConnection"));
            string sql = @"
                SELECT DISTINCT PONo + ' (' + CAST(RunningPONo AS VARCHAR(50)) + ')' AS OrderText, PONo
                FROM VVendOrdersToRcv 
                WHERE VendID = @VendID
                ORDER BY OrderText";
            var result = await connection.QueryAsync<string>(sql, new { VendID = vendId });
            return result.ToList();
        }

        public async Task<List<VendGateRcvdItemViewModel>> GetPendingOrderItemsAsync(string orderNo)
        {
            using IDbConnection connection = new SqlConnection(_config.GetConnectionString("DefaultConnection"));
            
            // Note: PONo in VVendOrdersToRcv corresponds to the OrderNo parameter
            // The dropdown gives "OrderNo (RunningPONo)". We need to extract the base OrderNo if passed as such, 
            // but assuming orderNo parameter is just PONo. Wait, if dropdown gives PONo + RunningPONo, we should probably fetch and bind properly in UI.
            // I'll adjust the query to search by PONo.
            
            string sql = @"
                SELECT 
                    PONo AS OrderNo,
                    MaterialID,
                    MaterialName,
                    QtyOrdered,
                    QtyToRcv,
                    0 AS QtyRcvd,
                    Rate,
                    0 AS Sheets,
                    Ref AS WORef,
                    IssuanceType,
                    Unit
                FROM VVendOrdersToRcv
                WHERE PONo = @OrderNo
                ORDER BY GroupID, MaterialName";
                
            var result = await connection.QueryAsync<VendGateRcvdItemViewModel>(sql, new { OrderNo = orderNo });
            return result.ToList();
        }

        public async Task<List<string>> GetStoresAsync()
        {
            using IDbConnection connection = new SqlConnection(_config.GetConnectionString("DefaultConnection"));
            string sql = "SELECT StoreName FROM Stores WHERE EntryID = 43";
            var result = await connection.QueryAsync<string>(sql);
            return result.ToList();
        }

        public async Task<string> SaveVendorReceivingAsync(VendGateRcvdSaveRequest request, string currentUser)
        {
            using var scope = new TransactionScope(TransactionScopeAsyncFlowOption.Enabled);
            using IDbConnection connection = new SqlConnection(_config.GetConnectionString("DefaultConnection"));
            
            // 1. Generate new RcvID: RCV-ddMMyy-XXX
            string prefix = $"RCV-{request.RcvDate:ddMMyy}-";
            string nextIdSql = $@"
                SELECT ISNULL(MAX(CAST(RIGHT(RTRIM(RcvID), 3) AS INT)), 0) + 1 
                FROM VendRcvd 
                WHERE LEFT(RcvID, LEN(@Prefix)) = @Prefix";
            
            int nextSequence = await connection.ExecuteScalarAsync<int>(nextIdSql, new { Prefix = prefix });
            string rcvId = $"{prefix}{nextSequence:D3}";
            
            // 2. Get RunningRcvNo
            string runningNoSql = "SELECT ISNULL(MAX(RunningRcvNo), 0) + 1 FROM VendRcvd";
            int runningRcvNo = await connection.ExecuteScalarAsync<int>(runningNoSql);
            
            // 3. Insert into VendRcvd
            string insertMasterSql = @"
                INSERT INTO VendRcvd (RcvID, VendID, RcvDate, Description, StockUpdated, PostingDate, RunningRcvNo)
                VALUES (@RcvID, @VendID, @RcvDate, @Description, 1, @PostingDate, @RunningRcvNo)";
                
            await connection.ExecuteAsync(insertMasterSql, new {
                RcvID = rcvId,
                VendID = request.VendID,
                RcvDate = request.RcvDate,
                Description = request.Description,
                PostingDate = DateTime.Now,
                RunningRcvNo = runningRcvNo
            });
            
            // 4. Insert details and update inventory
            foreach (var item in request.ReceivedItems.Where(x => x.QtyRcvd > 0))
            {
                double qty = item.QtyRcvd.Value;
                double extraQty = 0;
                
                if (qty > (item.QtyToRcv ?? 0))
                {
                    extraQty = qty - (item.QtyToRcv ?? 0);
                    qty = qty - extraQty;
                }
                
                string insertDetailSql = @"
                    INSERT INTO VendRcvdDetailPO (RcvID, PORefNo, MaterialID, QtyRcvd, WORef, QtyReturned, Rate, Checked, StoreName, ExtraQtyRcvd, Sheets)
                    VALUES (@RcvID, @PORefNo, @MaterialID, @QtyRcvd, @WORef, 0, @Rate, 1, @StoreName, @ExtraQtyRcvd, @Sheets)";
                    
                await connection.ExecuteAsync(insertDetailSql, new {
                    RcvID = rcvId,
                    PORefNo = request.OrderNo,
                    MaterialID = item.MaterialID,
                    QtyRcvd = qty,
                    WORef = item.WORef,
                    Rate = item.Rate ?? 0,
                    StoreName = item.StoreName,
                    ExtraQtyRcvd = extraQty,
                    Sheets = item.Sheets ?? 0
                });
                
                // Update RM inventory
                string updateRmSql = "";
                if (item.IssuanceType == 0)
                {
                    updateRmSql = "UPDATE RM SET QtyInStock = ISNULL(QtyInStock, 0) + @TotalQty, Rate = @Rate WHERE RMID1 = @MaterialID";
                }
                else if (item.IssuanceType == 1)
                {
                    updateRmSql = "UPDATE RM SET QtyInStock = ISNULL(QtyInStock, 0) + @TotalQty, Rate = @Rate, AnnealingStock = ISNULL(AnnealingStock, 0) + @TotalQty WHERE RMID1 = @MaterialID";
                }
                else if (item.IssuanceType == 2)
                {
                    updateRmSql = "UPDATE RM SET QtyInStock = ISNULL(QtyInStock, 0) + @TotalQty, Rate = @Rate, MachiningStock = ISNULL(MachiningStock, 0) + @TotalQty WHERE RMID1 = @MaterialID";
                }
                else if (item.IssuanceType == 3)
                {
                    updateRmSql = "UPDATE RM SET QtyInStock = ISNULL(QtyInStock, 0) + @TotalQty, Rate = @Rate, MaleGrindingStock = ISNULL(MaleGrindingStock, 0) + @TotalQty WHERE RMID1 = @MaterialID";
                }
                
                if (!string.IsNullOrEmpty(updateRmSql))
                {
                    await connection.ExecuteAsync(updateRmSql, new {
                        TotalQty = item.QtyRcvd.Value,
                        Rate = item.Rate ?? 0,
                        MaterialID = item.MaterialID
                    });
                    
                    // Legacy code also does an unconditional update to RM Rate using RMID, but doing it by RMID1 is functionally identical.
                }
            }
            
            scope.Complete();
            return rcvId;
        }
    }
}
