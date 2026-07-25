using Dapper;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Stock
{
    public class VendOrderDataAccess : IVendOrderDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<VendOrderDataAccess> _logger;

        public VendOrderDataAccess(IConfiguration configuration, ILogger<VendOrderDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                             ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        public async Task<IEnumerable<RMPOListViewModel>> GetRMPOListAsync(RMPOListSearchFilter filter)
        {
            try
            {
                using var connection = new SqlConnection(_connectionString);
                
                var conditions = new List<string>();
                var parameters = new DynamicParameters();

                if (!string.IsNullOrEmpty(filter.VendorID) && filter.VendorID != "0")
                {
                    conditions.Add("VendID = @VendorID");
                    parameters.Add("@VendorID", filter.VendorID);
                }

                if (filter.DateFrom.HasValue && filter.DateTo.HasValue)
                {
                    conditions.Add("DT BETWEEN @DateFrom AND @DateTo");
                    parameters.Add("@DateFrom", filter.DateFrom.Value);
                    parameters.Add("@DateTo", filter.DateTo.Value);
                }

                if (!string.IsNullOrEmpty(filter.MaterialID) && filter.MaterialID != "0")
                {
                    conditions.Add("OrderNo IN (SELECT OrderNo FROM VendOrderDetail WHERE RMID = @MaterialID)");
                    parameters.Add("@MaterialID", filter.MaterialID);
                }

                if (filter.RMGroupID.HasValue && filter.RMGroupID.Value != 0)
                {
                    conditions.Add("OrderNo IN (SELECT OrderNo FROM VVendOrderDetail WHERE GroupID = @RMGroupID)");
                    parameters.Add("@RMGroupID", filter.RMGroupID.Value);
                }

                if (filter.POType.HasValue && filter.POType.Value < 3)
                {
                    conditions.Add("IssuanceType = @POType");
                    parameters.Add("@POType", filter.POType.Value);
                }

                if (!string.IsNullOrEmpty(filter.OrderNo))
                {
                    conditions.Add("OrderNo = @OrderNo");
                    parameters.Add("@OrderNo", filter.OrderNo);
                }

                if (filter.ShowSamplePOs)
                {
                    conditions.Add("SampleOrder = 1");
                }
                else
                {
                    conditions.Add("(SampleOrder = 0 OR SampleOrder IS NULL)");
                }

                if (filter.ShowOpenPOsOnly)
                {
                    conditions.Add("RcvableOrderNo IS NOT NULL");
                }

                if (filter.ShowLateOrdersOnly)
                {
                    conditions.Add("DeliveryDT <= GETDATE()");
                }

                string whereClause = conditions.Count > 0 ? "WHERE " + string.Join(" AND ", conditions) : "";

                string sql = $@"
                    SELECT DT, VendID, AccTitle, AttnPerson, PaymentTerms, 
                           Department, DeliveryDT, Remarks, OrderNo, ReqNo, Final, 
                           RunningPONo, IssuanceType, CAST(CASE WHEN PDFAttachment IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS HasPDF, 
                           TotalQtyOrdered, SampleOrder, RcvableOrderNo, TotalQtyRcvd
                    FROM VVendOrdersList 
                    {whereClause} 
                    ORDER BY DT ASC";

                // Since we only need to know if PDF is attached, we can just use the HasPDF bit.
                // But wait, the ViewModel has byte[] PDFAttachment. Let's just fetch it normally.

                string finalSql = $@"
                    SELECT DT, VendID, AccTitle, AttnPerson, PaymentTerms, 
                           Department, DeliveryDT, Remarks, OrderNo, ReqNo, Final, 
                           RunningPONo, IssuanceType, PDFAttachment, TotalQtyOrdered, SampleOrder, 
                           RcvableOrderNo, TotalQtyRcvd
                    FROM VVendOrdersList 
                    {whereClause} 
                    ORDER BY DT ASC";

                return await connection.QueryAsync<RMPOListViewModel>(finalSql, parameters);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting RM PO List");
                throw;
            }
        }

        public async Task UpdatePDFAttachmentAsync(string orderNo, byte[] pdfData)
        {
            try
            {
                using var connection = new SqlConnection(_connectionString);
                string sql = "UPDATE VendOrders SET PDFAttachment = @PdfData WHERE OrderNo = @OrderNo";
                await connection.ExecuteAsync(sql, new { PdfData = pdfData, OrderNo = orderNo });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating PDF attachment");
                throw;
            }
        }

        public async Task<string> GetNextOrderNoAsync()
        {
            try
            {
                using var connection = new SqlConnection(_connectionString);
                string sql = "SELECT ISNULL(MAX(RunningPONo), 0) + 1 FROM VendOrders";
                var nextNo = await connection.QueryFirstOrDefaultAsync<long>(sql);
                // Based on legacy logic for POrdNo generation, this should be formatted as needed, but for now just returning the running PO No as string.
                return nextNo.ToString();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting next Order No");
                throw;
            }
        }

        public async Task<VendOrderViewModel> GetVendOrderAsync(string orderNo)
        {
            var model = new VendOrderViewModel();

            try
            {
                string sqlMaster = @"
                    SELECT *
                    FROM VendOrders
                    WHERE OrderNo = @orderNo;
                ";

                using var connection = new SqlConnection(_connectionString);
                var result = await connection.QueryAsync<VendOrderViewModel>(sqlMaster, new { orderNo });
                model = result.FirstOrDefault();

                if (model != null)
                {
                    string sqlDetails = @"
                        SELECT 
                            d.EntryID, d.OrderNo, d.RMID, d.Rate, d.Qty, d.ForRMID, d.CCItem, d.ItemRemarks,
                            r.RMName, r.Unit,
                            'Raw Material' AS RMType
                        FROM VendOrderDetail d
                        LEFT JOIN RM r ON r.RMID1 = d.RMID
                        WHERE d.OrderNo = @orderNo;
                    ";
                    var details = await connection.QueryAsync<VendOrderDetailViewModel>(sqlDetails, new { orderNo });
                    model.OrderDetails = details.ToList();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching VendOrder");
                throw;
            }

            return model ?? new VendOrderViewModel();
        }

        public async Task<string> InsertVendOrderAsync(VendOrderViewModel model)
        {
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();
            using var transaction = connection.BeginTransaction();

            try
            {
                // Generate RunningPONo
                string sqlRunningPo = "SELECT ISNULL(MAX(RunningPONo), 0) + 1 FROM VendOrders";
                model.RunningPONo = await connection.QueryFirstOrDefaultAsync<int>(sqlRunningPo, null, transaction);

                if (string.IsNullOrWhiteSpace(model.OrderNo))
                {
                    string monthYear = model.DT?.ToString("MMyy") ?? DateTime.Now.ToString("MMyy");
                    string sqlMaxOrder = @"
                        SELECT ISNULL(MAX(CAST(LEFT(OrderNo, CHARINDEX('/', OrderNo) - 1) AS INT)), 0) 
                        FROM VendOrders 
                        WHERE RIGHT(OrderNo, 4) = @MonthYear AND CHARINDEX('/', OrderNo) > 0";
                    
                    int lastSeq = await connection.QueryFirstOrDefaultAsync<int>(sqlMaxOrder, new { MonthYear = monthYear }, transaction);
                    model.OrderNo = $"{lastSeq + 1}/{monthYear}";
                }

                model.DT = model.DT?.Date;
                model.DeliveryDT = model.DeliveryDT?.Date;

                string sqlInsertMaster = @"
                    INSERT INTO VendOrders(
                        OrderNo, VendID, DT, AttnPerson, PaymentTerms, Department, DeliveryDT, Remarks, ReqNo,
                        VeryUrgent, RunningPONo, IssuanceType, Ref, Payments, Shipment, Instructions, Notes, Currency, SampleProvided,
                        SampleProvidedDescription, DrawingProvided, DrawingProvidedDescription, SteelProvided, SteelProvidedDescription,
                        POType, CustomerPONo, SteelType_RefID, SampleOrder, Order_Collection_By
                    ) VALUES (
                        @OrderNo, @VendID, @DT, @AttnPerson, @PaymentTerms, @Department, @DeliveryDT, @Remarks, @ReqNo,
                        @VeryUrgent, @RunningPONo, @IssuanceType, @Ref, @Payments, @Shipment, @Instructions, @Notes, @Currency, @SampleProvided,
                        @SampleProvidedDescription, @DrawingProvided, @DrawingProvidedDescription, @SteelProvided, @SteelProvidedDescription,
                        @POType, @CustomerPONo, @SteelType_RefID, @SampleOrder, @Order_Collection_By
                    );
                ";
                await connection.ExecuteAsync(sqlInsertMaster, model, transaction);

                if (model.OrderDetails != null && model.OrderDetails.Any())
                {
                    string sqlInsertDetail = @"
                        INSERT INTO VendOrderDetail (OrderNo, RMID, Rate, Qty, ForRMID, CCItem, ItemRemarks)
                        VALUES (@OrderNo, @RMID, @Rate, @Qty, @ForRMID, @CCItem, @ItemRemarks);
                    ";

                    foreach (var detail in model.OrderDetails)
                    {
                        detail.OrderNo = model.OrderNo;
                        await connection.ExecuteAsync(sqlInsertDetail, detail, transaction);
                    }
                }

                transaction.Commit();
                return model.OrderNo;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                _logger.LogError(ex, "Error inserting VendOrder");
                throw;
            }
        }

        public async Task<bool> UpdateVendOrderAsync(VendOrderViewModel model)
        {
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();
            using var transaction = connection.BeginTransaction();

            try
            {
                model.DT = model.DT?.Date;
                model.DeliveryDT = model.DeliveryDT?.Date;

                string sqlUpdateMaster = @"
                    UPDATE VendOrders SET 
                        DT = @DT, AttnPerson = @AttnPerson, PaymentTerms = @PaymentTerms, Department = @Department, 
                        DeliveryDT = @DeliveryDT, Remarks = @Remarks, ReqNo = @ReqNo, VeryUrgent = @VeryUrgent, 
                        IssuanceType = @IssuanceType, Ref = @Ref, Payments = @Payments, Shipment = @Shipment, 
                        Instructions = @Instructions, Notes = @Notes, Currency = @Currency, SampleProvided = @SampleProvided, 
                        SampleProvidedDescription = @SampleProvidedDescription, DrawingProvided = @DrawingProvided, 
                        DrawingProvidedDescription = @DrawingProvidedDescription, SteelProvided = @SteelProvided, 
                        SteelProvidedDescription = @SteelProvidedDescription, POType = @POType, CustomerPONo = @CustomerPONo, 
                        SteelType_RefID = @SteelType_RefID, SampleOrder = @SampleOrder, Order_Collection_By = @Order_Collection_By
                    WHERE OrderNo = @OrderNo;
                ";
                await connection.ExecuteAsync(sqlUpdateMaster, model, transaction);

                // Delete all details and re-insert
                string sqlDeleteDetails = "DELETE FROM VendOrderDetail WHERE OrderNo = @OrderNo";
                await connection.ExecuteAsync(sqlDeleteDetails, new { model.OrderNo }, transaction);

                if (model.OrderDetails != null && model.OrderDetails.Any())
                {
                    string sqlInsertDetail = @"
                        INSERT INTO VendOrderDetail (OrderNo, RMID, Rate, Qty, ForRMID, CCItem, ItemRemarks)
                        VALUES (@OrderNo, @RMID, @Rate, @Qty, @ForRMID, @CCItem, @ItemRemarks);
                    ";

                    foreach (var detail in model.OrderDetails)
                    {
                        detail.OrderNo = model.OrderNo;
                        await connection.ExecuteAsync(sqlInsertDetail, detail, transaction);
                    }
                }

                transaction.Commit();
                return true;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                _logger.LogError(ex, "Error updating VendOrder");
                throw;
            }
        }

        public async Task<bool> DeleteVendOrderAsync(string orderNo)
        {
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();
            using var transaction = connection.BeginTransaction();

            try
            {
                string sqlDeleteDetails = "DELETE FROM VendOrderDetail WHERE OrderNo = @orderNo";
                await connection.ExecuteAsync(sqlDeleteDetails, new { orderNo }, transaction);

                string sqlDeleteMaster = "DELETE FROM VendOrders WHERE OrderNo = @orderNo";
                await connection.ExecuteAsync(sqlDeleteMaster, new { orderNo }, transaction);

                transaction.Commit();
                return true;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                _logger.LogError(ex, "Error deleting VendOrder");
                throw;
            }
        }
    }
}
