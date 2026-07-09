using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Dapper;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;

namespace DataAccessLibrary.DAC.Export
{
    public class OrderItemListDataAccess : IOrderItemListDataAccess
    {
        private readonly string _connectionString;

        public OrderItemListDataAccess(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")!;
        }

        public async Task<List<OrderItemListViewModel>> GetOrderItemsAsync(
            DateTime dtFrom,
            DateTime dtTo,
            string? custCode,
            string? country,
            int? groupID,
            string? orderNo,
            bool noDelivery,
            bool both,
            bool urgent)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            var cond = new StringBuilder();

            // --- Primary date/urgency condition (mirrors legacy getCondition()) ---
            if (urgent)
            {
                cond.Append(" WHERE (VFOrderItems.DeliveryStatus=2 OR VFOrderItems.Remarks<>'')");
            }
            else if (both)
            {
                cond.Append($" WHERE (VFOrderItems.DeliveryStatus=0 OR (VFOrderItems.DeliveryDT BETWEEN @DtFrom AND @DtTo))");
            }
            else if (noDelivery)
            {
                cond.Append(" WHERE (VFOrderItems.DeliveryStatus=0 OR VFOrderItems.DeliveryDT='1900-01-01')");
            }
            else
            {
                cond.Append(" WHERE VFOrderItems.DeliveryDT BETWEEN @DtFrom AND @DtTo");
            }

            // Always: only balance items, in open orders, not stock
            cond.Append(" AND (Qty > ShippedQty)");
            cond.Append(" AND VFOrderItems.OrderNo IN (SELECT OrderNo FROM VUnshippedOrderList)");
            cond.Append(" AND VFOrderItems.CustCode NOT IN('Stock')");

            if (!string.IsNullOrEmpty(custCode))
                cond.Append(" AND VFOrderItems.CustCode = @CustCode");

            if (!string.IsNullOrEmpty(country))
                cond.Append(" AND VFOrderItems.Country = @Country");

            if (groupID.HasValue && groupID.Value > 0)
                cond.Append(" AND GroupID = @GroupID");

            if (!string.IsNullOrEmpty(orderNo))
                cond.Append(" AND VFOrderItems.OrderNo = @OrderNo");

            var sql = $@"
                SELECT 
                    VFOrderItems.ID,
                    VFOrderItems.CustCode,
                    VFOrderItems.Country,
                    VFOrderItems.OrderNo,
                    VFOrderItems.InternalRefNo,
                    VFOrderItems.CompItemCode,
                    VFOrderItems.ItemName,
                    VFOrderItems.DeliveryDT,
                    VFOrderItems.DeliveryStatus,
                    VFOrderItems.Qty,
                    VFOrderItems.ShippedQty,
                    VFOrderItems.Remarks,
                    VFOrderItems.GroupID,
                    ISNULL(TDisp.InnerQty, 0) AS InnerQty,
                    ISNULL(dbo.DisptachListNos(VFOrderItems.OrderNo, VFOrderItems.CompItemCode), '') AS DispatchListNos
                FROM VFOrderItems
                LEFT OUTER JOIN (
                    SELECT OrderNo, ItemCode, SUM(InnerQty) AS InnerQty
                    FROM VDispatchListDetail_Ex_Orderwise
                    WHERE Finalyzed = 0
                    GROUP BY OrderNo, ItemCode
                ) TDisp ON VFOrderItems.OrderNo = TDisp.OrderNo 
                       AND VFOrderItems.CompItemCode = TDisp.ItemCode
                {cond}
                ORDER BY VFOrderItems.DeliveryDT DESC";

            var result = await db.QueryAsync<OrderItemListViewModel>(sql, new
            {
                DtFrom = dtFrom,
                DtTo = dtTo,
                CustCode = custCode,
                Country = country,
                GroupID = groupID,
                OrderNo = orderNo
            });

            return result.AsList();
        }

        public async Task<List<string>> GetCountriesForCustomerAsync(string custCode)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql;
            if (string.IsNullOrEmpty(custCode))
                sql = "SELECT DISTINCT Country FROM ForeignCustomers ORDER BY Country";
            else
                sql = "SELECT DISTINCT Country FROM ForeignCustomers WHERE CustCode = @CustCode ORDER BY Country";

            var result = await db.QueryAsync<string>(sql, new { CustCode = custCode });
            return result.AsList();
        }

        public async Task<List<OrderNoLookupModel>> GetOrderNosForCustomerAsync(string custCode)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql;
            if (string.IsNullOrEmpty(custCode))
                sql = "SELECT OrderNo, ISNULL(InternalRefNo,'') AS InternalRefNo FROM FCustomerOrders ORDER BY OrderNo";
            else
                sql = "SELECT OrderNo, ISNULL(InternalRefNo,'') AS InternalRefNo FROM FCustomerOrders WHERE CustCode = @CustCode ORDER BY OrderNo";

            var result = await db.QueryAsync<OrderNoLookupModel>(sql, new { CustCode = custCode });
            return result.AsList();
        }

        public async Task<List<OrderItemGroupLookupModel>> GetItemGroupsAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT ID, Description FROM ItemGroups ORDER BY Description";
            var result = await db.QueryAsync<OrderItemGroupLookupModel>(sql);
            return result.AsList();
        }

        public async Task<bool> UpdateDeliveryDateAsync(
            int id,
            DateTime newDeliveryDT,
            int prevDeliveryStatus,
            DateTime? prevDeliveryDT,
            string remarks,
            string userName,
            string machineName)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            db.Open();
            using var tran = db.BeginTransaction();
            try
            {
                // Update the delivery date on FOrderItems; keep existing status if already set, else move to 1
                const string updateSql = @"
                    UPDATE FOrderItems
                    SET DeliveryDT = @DeliveryDT,
                        DeliveryStatus = CASE WHEN DeliveryStatus = 0 THEN 1 ELSE DeliveryStatus END
                    WHERE ID = @ID";

                await db.ExecuteAsync(updateSql, new { DeliveryDT = newDeliveryDT, ID = id }, tran);

                // Insert revision history record
                const string insertSql = @"
                    INSERT INTO FOrderItems_DeliveryRevisions
                        (FOI_RefID, DeliveryStatus, DeliveryDT_Previous, DeliveryDT, Remarks, UserName, MachineName)
                    VALUES
                        (@ID, @PrevStatus, @PrevDT, @DeliveryDT, @Remarks, @UserName, @MachineName)";

                await db.ExecuteAsync(insertSql, new
                {
                    ID = id,
                    PrevStatus = prevDeliveryStatus,
                    PrevDT = prevDeliveryDT,
                    DeliveryDT = newDeliveryDT,
                    Remarks = remarks,
                    UserName = userName,
                    MachineName = machineName
                }, tran);

                tran.Commit();
                return true;
            }
            catch
            {
                tran.Rollback();
                throw;
            }
        }
    }
}
