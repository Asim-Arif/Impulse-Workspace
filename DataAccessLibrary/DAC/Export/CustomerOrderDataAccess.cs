using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Export;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Export
{
    public class CustomerOrderDataAccess : ICustomerOrderDataAccess
    {
        private readonly string _connectionString;

        public CustomerOrderDataAccess(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
        }

        public async Task<List<GenericDropDownModel>> GetCompaniesAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT EntryID AS DropDownValue_ID, CompanyName AS DropDownValue_Description FROM Companies ORDER BY CompanyName";
            return (await db.QueryAsync<GenericDropDownModel>(sql)).ToList();
        }

        public async Task<List<CustomerLookupModel>> GetCustomersAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT CustCode, Name FROM ForeignCustomers WHERE Active = 1 ORDER BY Name";
            return (await db.QueryAsync<CustomerLookupModel>(sql)).ToList();
        }

        public async Task<List<GenericDropDownModel>> GetTradeTermsAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT ID AS DropDownValue_ID, Description AS DropDownValue_Description FROM TradeTerms ORDER BY ID";
            return (await db.QueryAsync<GenericDropDownModel>(sql)).ToList();
        }

        public async Task<List<GenericDropDownModel>> GetPaymentTermsAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT ID AS DropDownValue_ID, Description AS DropDownValue_Description FROM PaymentTerms ORDER BY Description";
            return (await db.QueryAsync<GenericDropDownModel>(sql)).ToList();
        }

        public async Task<List<GenericDropDownModel>> GetShippingMethodsAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT EntryID AS DropDownValue_ID, Description AS DropDownValue_Description FROM ShippingMethods ORDER BY EntryID";
            return (await db.QueryAsync<GenericDropDownModel>(sql)).ToList();
        }

        public async Task<List<GenericDropDownModel>> GetQualitiesAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT EntryID AS DropDownValue_ID, Quality AS DropDownValue_Description FROM FCO_Qualities ORDER BY Quality";
            return (await db.QueryAsync<GenericDropDownModel>(sql)).ToList();
        }

        public async Task<List<GenericDropDownModel>> GetFinishingQualitiesAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT ID AS DropDownValue_ID, Description AS DropDownValue_Description FROM ItemsFinQuality ORDER BY Description";
            return (await db.QueryAsync<GenericDropDownModel>(sql)).ToList();
        }

        public async Task<List<string>> GetCountriesForCustomerAsync(string custCode)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT DISTINCT Country FROM ForeignCustomers WHERE CustCode = @CustCode";
            return (await db.QueryAsync<string>(sql, new { CustCode = custCode })).ToList();
        }

        public async Task<string> GetCurrencyForCustomerAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT TOP 1 Curr FROM ForeignCustomers WHERE CustCode = @CustCode AND Country = @Country";
            var curr = await db.QueryFirstOrDefaultAsync<string>(sql, new { CustCode = custCode, Country = country });
            return curr ?? string.Empty;
        }

        public async Task<ForeignCustomerViewModel?> GetCustomerDefaultsAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT TradeTerms, PaymentTerms, FinishingQuality, Stamps, RcvdVia 
                FROM ForeignCustomers 
                WHERE CustCode = @CustCode AND Country = @Country";
            return await db.QueryFirstOrDefaultAsync<ForeignCustomerViewModel>(sql, new { CustCode = custCode, Country = country });
        }

        public async Task<List<CustomerOrderItemViewModel>> GetCustomerCatalogItemsAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT CompItemID AS CompItemCode, ItemID AS ItemCode, Description, ExWorks AS Price, 0 AS Weight, Unit 
                FROM VFCustomerCatalog 
                WHERE CustCode = @CustCode AND Country = @Country AND InActive = 0";
            return (await db.QueryAsync<CustomerOrderItemViewModel>(sql, new { CustCode = custCode, Country = country })).ToList();
        }

        public async Task<decimal> GetItemCustomRateAsync(string custCode, string country, string compItemId, string finishingField)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            // Dynamic column name whitelisting to prevent SQL injection
            string field = finishingField == "First_Inspection_Rate" ? "First_Inspection_Rate" : "Temper_Rate";
            string sql = $@"
                SELECT TOP 1 {field} 
                FROM FCustomerCatalog 
                WHERE CustCode = @CustCode AND Country = @Country AND CompItemID = @CompItemId";
            return await db.QueryFirstOrDefaultAsync<decimal>(sql, new { CustCode = custCode, Country = country, CompItemId = compItemId });
        }

        public async Task<decimal> GetItemWeightAsync(string itemId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT TOP 1 FinishedWeight FROM Items WHERE ItemID = @ItemId";
            return await db.QueryFirstOrDefaultAsync<decimal>(sql, new { ItemId = itemId });
        }

        public async Task<List<string>> GetQuotationNumbersAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT DISTINCT QuotationNo FROM CustomerQuotations WHERE CustCode = @CustCode AND Country = @Country ORDER BY QuotationNo";
            return (await db.QueryAsync<string>(sql, new { CustCode = custCode, Country = country })).ToList();
        }

        public async Task<List<CustomerOrderItemViewModel>> GetQuotationItemsAsync(string quotationNo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT CompItemID AS CompItemCode, ItemID AS ItemCode, ItemName AS Description, Price, Qty, Unit 
                FROM VCustomerQuotations 
                WHERE QuotationNo = @QuotationNo";
            return (await db.QueryAsync<CustomerOrderItemViewModel>(sql, new { QuotationNo = quotationNo })).ToList();
        }

        public async Task<CustomerOrderViewModel?> GetOrderAsync(string orderNo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string orderSql = "SELECT * FROM FCustomerOrders WHERE OrderNo = @OrderNo";
            var order = await db.QueryFirstOrDefaultAsync<CustomerOrderViewModel>(orderSql, new { OrderNo = orderNo });
            if (order == null) return null;

            const string itemsSql = @"
                SELECT ID, OrderNo, ItemCode, CompItemCode, Price, Qty, InvQty, CustomPrice, SortNo, DeliveryDT, 
                       StampsItem AS Stamps, Quality, DeliveryStatus, Remarks, Weight, Item_Finishing_Type, IW_OrderNo, IW_BatchNo, 
                       ItemName, Description, Unit, Item_Finishing_Type_Text
                FROM VrptOrders 
                WHERE OrderNo = @OrderNo 
                ORDER BY SortNo";
            order.OrderItems = (await db.QueryAsync<CustomerOrderItemViewModel>(itemsSql, new { OrderNo = orderNo })).ToList();

            return order;
        }

        public async Task<bool> SaveOrderAsync(CustomerOrderViewModel order, List<int> deletedItemIds)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            db.Open();
            using var trans = db.BeginTransaction();
            try
            {
                // Check if existing
                const string checkSql = "SELECT COUNT(*) FROM FCustomerOrders WHERE OrderNo = @OrderNo";
                int count = await db.QueryFirstOrDefaultAsync<int>(checkSql, new { OrderNo = order.OrderNo }, trans);

                if (count > 0)
                {
                    // Update
                    const string updateOrderSql = @"
                        UPDATE FCustomerOrders 
                        SET DT = @DT, TradeTerms = @TradeTerms, PartialShipment = @PartialShipment, 
                            PaymentTerms = @PaymentTerms, TransShipment = @TransShipment, Packaging = @Packaging, 
                            DeliveryDT = @DeliveryDT, CompanyRefID = @CompanyRefID, StampDT = @StampDT, 
                            Quality = @Quality, InternalRefNo = @InternalRefNo, OrderRcvdVia = @OrderRcvdVia, 
                            OrderType = @OrderType, ShippingMode = @ShippingMode, OrderRevisionNo = @OrderRevisionNo, 
                            OrderRevisionDT = @OrderRevisionDT, Packaging_Weight = @Packaging_Weight 
                        WHERE OrderNo = @OrderNo";
                    await db.ExecuteAsync(updateOrderSql, order, trans);
                }
                else
                {
                    // Insert
                    const string insertOrderSql = @"
                        INSERT INTO FCustomerOrders (CustCode, Country, OrderNo, DT, TradeTerms, PartialShipment, 
                                                    PaymentTerms, TransShipment, Packaging, DeliveryDT, CompanyRefID, 
                                                    StampDT, Quality, InternalRefNo, OrderRcvdVia, OrderType, 
                                                    ShippingMode, OrderRevisionNo, OrderRevisionDT, Packaging_Weight) 
                        VALUES (@CustCode, @Country, @OrderNo, @DT, @TradeTerms, @PartialShipment, 
                                @PaymentTerms, @TransShipment, @Packaging, @DeliveryDT, @CompanyRefID, 
                                @StampDT, @Quality, @InternalRefNo, @OrderRcvdVia, @OrderType, 
                                @ShippingMode, @OrderRevisionNo, @OrderRevisionDT, @Packaging_Weight)";
                    await db.ExecuteAsync(insertOrderSql, order, trans);
                }

                // Delete deleted items
                if (deletedItemIds != null && deletedItemIds.Any())
                {
                    const string deleteItemsSql = "DELETE FROM FOrderItems WHERE ID IN @Ids";
                    await db.ExecuteAsync(deleteItemsSql, new { Ids = deletedItemIds }, trans);
                }

                // Save Item lines
                foreach (var item in order.OrderItems)
                {
                    item.OrderNo = order.OrderNo;
                    if (item.ID == 0)
                    {
                        // Insert
                        const string insertItemSql = @"
                            INSERT INTO FOrderItems (OrderNo, ItemCode, CompItemCode, Price, Qty, InvQty, 
                                                     CustomPrice, SortNo, DeliveryDT, Quality, Stamps, 
                                                     DeliveryStatus, Remarks, Weight, Item_Finishing_Type, 
                                                     IW_OrderNo, IW_BatchNo) 
                            VALUES (@OrderNo, @ItemCode, @CompItemCode, @Price, @Qty, @Qty, 
                                    @CustomPrice, @SortNo, @DeliveryDT, @Quality, @Stamps, 
                                    @DeliveryStatus, @Remarks, @Weight, @Item_Finishing_Type, 
                                    @IW_OrderNo, @IW_BatchNo)";
                        await db.ExecuteAsync(insertItemSql, item, trans);
                    }
                    else
                    {
                        // Update
                        const string updateItemSql = @"
                            UPDATE FOrderItems 
                            SET Price = @Price, Qty = @Qty, SortNo = @SortNo, DeliveryDT = @DeliveryDT, 
                                Quality = @Quality, Stamps = @Stamps, DeliveryStatus = @DeliveryStatus, 
                                Remarks = @Remarks, Weight = @Weight, Item_Finishing_Type = @Item_Finishing_Type, 
                                IW_OrderNo = @IW_OrderNo, IW_BatchNo = @IW_BatchNo 
                            WHERE ID = @ID";
                        await db.ExecuteAsync(updateItemSql, item, trans);
                    }

                    // Update price in catalog as per VB6 logic
                    if (item.Item_Finishing_Type == 0)
                    {
                        const string updateCatSql = @"
                            UPDATE FCustomerCatalog 
                            SET ExWorks = @Price 
                            WHERE CustCode = @CustCode AND Country = @Country AND CompItemID = @CompItemCode";
                        int affected = await db.ExecuteAsync(updateCatSql, new { Price = item.Price, CustCode = order.CustCode, Country = order.Country, CompItemCode = item.CompItemCode }, trans);
                        
                        if (affected == 0)
                        {
                            const string insertCatSql = @"
                                INSERT INTO FCustomerCatalog (CustCode, Country, CompItemID, ItemID, Description, ExWorks) 
                                VALUES (@CustCode, @Country, @CompItemCode, @ItemCode, @Description, @Price)";
                            await db.ExecuteAsync(insertCatSql, new { CustCode = order.CustCode, Country = order.Country, CompItemCode = item.CompItemCode, ItemCode = item.ItemCode, Description = item.Description, Price = item.Price }, trans);
                        }
                    }
                }

                trans.Commit();
                return true;
            }
            catch (Exception)
            {
                trans.Rollback();
                throw;
            }
        }

        public async Task<string> GetNextInternalRefNoAsync(int companyId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            // Default max query logic matching VB6 GetInternalRefNo
            const string sql = "SELECT MAX(CAST(InternalRefNo AS INT)) AS MaxNo FROM FCustomerOrders WHERE ISNUMERIC(InternalRefNo) = 1";
            int? maxResult = await db.QueryFirstOrDefaultAsync<int?>(sql);
            int max;
            if (maxResult == null)
            {
                max = 596; // default fallback matching InputBox default in VB6
            }
            else
            {
                max=maxResult.Value + 1;
            }
            return max.ToString();
        }

        public async Task<bool> IsOrderNoExistsAsync(string orderNo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT COUNT(*) AS TotalNos FROM FCustomerOrders WHERE OrderNo = @OrderNo";
            return await db.QueryFirstOrDefaultAsync<int>(sql, new { OrderNo = orderNo }) > 0;
        }

        public async Task<bool> IsOrderUsedInProformaAsync(string orderNo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT COUNT(*) AS TotalNos FROM FProformaOrders WHERE OrderEntryID IN (SELECT ID FROM FOrderItems WHERE OrderNo = @OrderNo)";
            return await db.QueryFirstOrDefaultAsync<int>(sql, new { OrderNo = orderNo }) > 0;
        }

        // ------------------------------------------------------------------ //
        //  Excel Import Operations
        // ------------------------------------------------------------------ //

        public async Task<bool> ClearExcelImportTableAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            await db.ExecuteAsync("DELETE FROM Import_Excel_CustomerOrder");
            return true;
        }

        public async Task<bool> InsertExcelImportRowAsync(string itemId, string itemName, decimal qty, string unit, decimal rate, decimal weight, string steelType, string udms)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                INSERT INTO Import_Excel_CustomerOrder (ItemID, ItemName, OrderQty, Unit, Rate, Weight, Steel_Type, UDMS) 
                VALUES (@ItemID, @ItemName, @OrderQty, @Unit, @Rate, @Weight, @SteelType, @UDMS)";
            await db.ExecuteAsync(sql, new { ItemID = itemId, ItemName = itemName, OrderQty = qty, Unit = unit, Rate = rate, Weight = weight, SteelType = steelType, UDMS = udms });
            return true;
        }

        public async Task<bool> ProcessExcelImportAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            db.Open();
            using var trans = db.BeginTransaction();
            try
            {
                await db.ExecuteAsync(@"
                    UPDATE Import_Excel_CustomerOrder 
                    SET CodeInCompanyCatalog = FCustomerCatalog.CompItemID, 
                        CodeInCustomerCatalog = FCustomerCatalog.ItemID 
                    FROM Import_Excel_CustomerOrder 
                    INNER JOIN FCustomerCatalog ON FCustomerCatalog.ItemID = Import_Excel_CustomerOrder.ItemID 
                    WHERE CustCode = @CustCode AND Country = @Country", new { CustCode = custCode, Country = country }, trans);

                await db.ExecuteAsync(@"
                    UPDATE Import_Excel_CustomerOrder 
                    SET CodeInCompanyCatalog = FCustomerCatalog.CompItemID, 
                        CodeInCustomerCatalog = FCustomerCatalog.ItemID, 
                        FromOtherCustomer = 1 
                    FROM Import_Excel_CustomerOrder 
                    INNER JOIN FCustomerCatalog ON FCustomerCatalog.ItemID = Import_Excel_CustomerOrder.ItemID 
                    WHERE CodeInCompanyCatalog IS NULL", null, trans);

                await db.ExecuteAsync(@"
                    UPDATE Import_Excel_CustomerOrder 
                    SET CodeInCompanyCatalog = Items.ItemID 
                    FROM Import_Excel_CustomerOrder 
                    INNER JOIN Items ON Items.ItemID = Import_Excel_CustomerOrder.ItemID 
                    WHERE Import_Excel_CustomerOrder.CodeInCompanyCatalog IS NULL", null, trans);

                // Insert into catalog where FromOtherCustomer = 1
                await db.ExecuteAsync(@"
                    INSERT INTO FCustomerCatalog (CustCode, Country, ItemID, CompItemID, Unit, Description, ExWorks) 
                    SELECT @CustCode, @Country, CodeInCustomerCatalog, CodeInCompanyCatalog, 'Pcs', '', 0 
                    FROM Import_Excel_CustomerOrder 
                    WHERE FromOtherCustomer = 1", new { CustCode = custCode, Country = country }, trans);

                trans.Commit();
                return true;
            }
            catch (Exception)
            {
                trans.Rollback();
                throw;
            }
        }

        public async Task<List<CustomerOrderItemViewModel>> GetExcelImportErrorsAsync(bool compCatalogMissing)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = compCatalogMissing 
                ? "SELECT ItemID AS ItemCode, ItemName AS Description, OrderQty AS Qty FROM Import_Excel_CustomerOrder WHERE CodeInCompanyCatalog IS NULL ORDER BY EntryID"
                : "SELECT ItemID AS ItemCode, ItemName AS Description, OrderQty AS Qty FROM Import_Excel_CustomerOrder WHERE CodeInCompanyCatalog IS NOT NULL AND CodeInCustomerCatalog IS NULL ORDER BY EntryID";
            return (await db.QueryAsync<CustomerOrderItemViewModel>(sql)).ToList();
        }

        public async Task<List<CustomerOrderItemViewModel>> GetExcelImportOkAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT Import_Excel_CustomerOrder.CodeInCompanyCatalog AS CompItemCode, 
                       FCustomerCatalog.ItemID AS ItemCode, 
                       Import_Excel_CustomerOrder.ItemName AS Description, 
                       Import_Excel_CustomerOrder.OrderQty AS Qty, 
                       Import_Excel_CustomerOrder.Rate AS Price, 
                       Import_Excel_CustomerOrder.Weight 
                FROM Import_Excel_CustomerOrder 
                INNER JOIN FCustomerCatalog ON Import_Excel_CustomerOrder.CodeInCompanyCatalog = FCustomerCatalog.CompItemID 
                WHERE CustCode = @CustCode AND Country = @Country 
                ORDER BY Import_Excel_CustomerOrder.EntryID";
            return (await db.QueryAsync<CustomerOrderItemViewModel>(sql, new { CustCode = custCode, Country = country })).ToList();
        }

        public async Task<bool> DeleteCustomerOrderAsync(string orderNo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string checkSql = "SELECT COUNT(*) FROM FProformaOrders WHERE OrderEntryID IN (SELECT ID FROM FOrderItems WHERE OrderNo = @OrderNo)";
            int count = await db.QueryFirstOrDefaultAsync<int>(checkSql, new { OrderNo = orderNo });
            if (count > 0)
            {
                return false;
            }

            const string deleteSql = "DELETE FROM FCustomerOrders WHERE OrderNo = @OrderNo";
            await db.ExecuteAsync(deleteSql, new { OrderNo = orderNo });
            return true;
        }

        public async Task<List<CustomerOrderListItemModel>> GetOrderListAsync(
            DateTime dtFrom, 
            DateTime dtTo, 
            string custCode, 
            string country, 
            int companyRefID, 
            string orderType, 
            int statusFilter, 
            bool filterByDeliveryDT, 
            int viewType)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            var parameters = new DynamicParameters();
            parameters.Add("@DtFrom", dtFrom);
            parameters.Add("@DtTo", dtTo);

            string dateField = filterByDeliveryDT ? "DeliveryDT" : "DT";
            string cond = $"WHERE {dateField} BETWEEN @DtFrom AND @DtTo";

            if (custCode != "0" && !string.IsNullOrEmpty(custCode))
            {
                cond += " AND CustCode = @CustCode";
                parameters.Add("@CustCode", custCode);
            }

            if (!string.IsNullOrEmpty(country) && country != "<All Countries>")
            {
                cond += " AND Country = @Country";
                parameters.Add("@Country", country);
            }

            if (companyRefID > 0)
            {
                cond += " AND CompanyRefID = @CompanyRefID";
                parameters.Add("@CompanyRefID", companyRefID);
            }

            if (!string.IsNullOrEmpty(orderType) && orderType != "<All Types>")
            {
                cond += " AND OrderType = @OrderType";
                parameters.Add("@OrderType", orderType);
            }

            if (viewType == 1)
            {
                cond += " AND CustCode NOT IN ('Parts','Stock')";
            }
            else if (viewType == 2)
            {
                cond += " AND CustCode IN ('Parts','Stock')";
            }

            string statusCond = "";
            if (statusFilter == 0) // Unshipped
            {
                statusCond = " (UnShippedOrderNo IS NOT NULL)";
            }
            else if (statusFilter == 1) // Shipped
            {
                statusCond = " (UnShippedOrderNo IS NULL AND Cancelled = 0)";
            }
            else if (statusFilter == 2) // Cancelled
            {
                statusCond = " Cancelled = 1";
            }
            else if (statusFilter == 3) // Hold
            {
                statusCond = " Cancelled = 2";
            }

            if (!string.IsNullOrEmpty(statusCond))
            {
                cond += " AND " + statusCond;
            }

            string sql = $"SELECT * FROM VFOrderList {cond} ORDER BY DT DESC";
            var list = (await db.QueryAsync<CustomerOrderListItemModel>(sql, parameters)).ToList();

            foreach (var item in list)
            {
                if (item.TotalShippedQty == 0 || item.TotalShippedQty == null)
                {
                    item.OrderStatus = "UnShipped";
                }
                else if (item.TotalOrderQty > item.TotalShippedQty)
                {
                    item.OrderStatus = "Partialy Shipped";
                }
                else
                {
                    item.OrderStatus = "Shipped";
                }

                if (item.Cancelled == 1)
                {
                    item.OrderStatus = "Cancelled";
                }
                else if (item.Cancelled == 2)
                {
                    item.OrderStatus = "Hold";
                }
                else
                {
                    if (item.UnshippedOrderNo == null)
                    {
                        item.OrderStatus = "Shipped";
                    }
                    else
                    {
                        item.OrderStatus = "UnShipped";
                    }
                }

                if (item.OrderPlanApproved == true)
                {
                    item.PurchasePlanStatus = "Approved";
                }
                else if (item.TotalPlannedQty > 0)
                {
                    item.PurchasePlanStatus = "Planned";
                }
                else
                {
                    item.PurchasePlanStatus = "-";
                }
            }

            return list;
        }

        public async Task<List<string>> GetAllCountriesAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT DISTINCT Country FROM ForeignCustomers WHERE Country IS NOT NULL AND Country <> '' ORDER BY Country";
            return (await db.QueryAsync<string>(sql)).ToList();
        }

        public async Task<bool> UpdateOrderFinalStatusAsync(
            string orderNo, 
            string custCode, 
            string country, 
            int cancelledStatus, 
            string remarks, 
            string userName, 
            string machineName)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                IF EXISTS (SELECT 1 FROM FCustomerFinalOrders WHERE OrderNo = @OrderNo)
                BEGIN
                    UPDATE FCustomerFinalOrders 
                    SET CustCode = @CustCode, Country = @Country, Cancelled = @Cancelled, Remarks = @Remarks, UserName = @UserName, MachineName = @MachineName, DTEntry = GETDATE()
                    WHERE OrderNo = @OrderNo
                END
                ELSE
                BEGIN
                    INSERT INTO FCustomerFinalOrders (CustCode, Country, OrderNo, UserName, MachineName, Cancelled, Remarks, DTEntry)
                    VALUES (@CustCode, @Country, @OrderNo, @UserName, @MachineName, @Cancelled, @Remarks, GETDATE())
                END";
            await db.ExecuteAsync(sql, new { OrderNo = orderNo, CustCode = custCode, Country = country, Cancelled = cancelledStatus, Remarks = remarks, UserName = userName, MachineName = machineName });
            return true;
        }

        public async Task<bool> DeleteOrderFinalStatusAsync(string orderNo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "DELETE FROM FCustomerFinalOrders WHERE OrderNo = @OrderNo";
            await db.ExecuteAsync(sql, new { OrderNo = orderNo });
            return true;
        }
    }
}
