using Dapper;
using DataAccessLibrary.Interface;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

using System.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Export
{
    public class CustomerItemBalancesDataAccess : ICustomerItemBalancesDataAccess
    {
        private readonly string _connectionString;

        public CustomerItemBalancesDataAccess(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection") 
                                ?? throw new System.Exception("Connection string 'DefaultConnection' not found.");
        }

        public async Task<List<CustomerItemBalancesViewModel>> GetCustomerItemBalancesAsync(string custCode)
        {
            string sql = @"
                SELECT 
                    ID, 
                    OrderNo, 
                    InternalRefNo, 
                    CompItemCode, 
                    ItemName, 
                    DeliveryDT, 
                    DeliveryStatus, 
                    Qty, 
                    ShippedQty
                FROM VFOrderItems 
                WHERE CustCode = @CustCode 
                  AND (Qty > ShippedQty) 
                  AND OrderNo IN (SELECT OrderNo FROM VUnshippedOrderList) 
                  AND CustCode NOT IN ('Stock')
                ORDER BY DeliveryDT DESC";

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var data = await db.QueryAsync<CustomerItemBalancesViewModel>(
                    sql,
                    new { CustCode = custCode });

                var dataList = data.ToList();
                // Compute BalQty
                foreach(var item in dataList)
                {
                    item.BalQty = item.Qty - item.ShippedQty;
                }

                return dataList;
            }
        }

        public async Task UpdateBalanceQuantityAsync(long id, decimal qtyDiff)
        {
            // The legacy system updates both FOrderItems and FProformaOrders
            string sqlItems = "UPDATE FOrderItems SET Qty = Qty + @QtyDiff WHERE ID = @ID";
            string sqlProforma = "UPDATE FProformaOrders SET Qty = Qty + @QtyDiff WHERE OrderEntryID = @ID";

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                await db.ExecuteAsync(sqlItems, new { ID = id, QtyDiff = qtyDiff });
                await db.ExecuteAsync(sqlProforma, new { ID = id, QtyDiff = qtyDiff });
            }
        }

        public async Task ToggleDeliveryStatusAsync(long id, int newStatus)
        {
            string sql = "UPDATE FOrderItems SET DeliveryStatus = @Status WHERE ID = @ID";
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                await db.ExecuteAsync(sql, new { ID = id, Status = newStatus });
            }
        }
    }
}
