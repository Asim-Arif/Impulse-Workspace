using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using Dapper;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;
using System;

namespace DataAccessLibrary.DAC.Export
{
    public class ArticlewiseShippedStatusDataAccess : IArticlewiseShippedStatusDataAccess
    {
        private readonly IConfiguration _config;

        public ArticlewiseShippedStatusDataAccess(IConfiguration config)
        {
            _config = config;
        }

        public async Task<IEnumerable<ArticleShippedInvoiceModel>> GetShippedInvoicesAsync(string itemCode, string custCode)
        {
            using IDbConnection connection = new SqlConnection(_config.GetConnectionString("DefaultConnection"));
            string sql = @"SELECT CustomInvoice, DT, GatePassDT, Qty, Price, OrderNo, OrderDT, DeliveryDTItem 
                           FROM VCustomInvoice 
                           WHERE CompItemCode = @ItemCode AND GatePassDT IS NOT NULL";
            
            if (!string.IsNullOrEmpty(custCode) && custCode != "0")
            {
                sql += " AND CustCode = @CustCode";
            }
            
            sql += " ORDER BY DT DESC";

            return await connection.QueryAsync<ArticleShippedInvoiceModel>(sql, new { ItemCode = itemCode, CustCode = custCode });
        }

        public async Task<IEnumerable<ArticleUnshippedOrderModel>> GetUnshippedOrdersAsync(string itemCode, string custCode)
        {
            using IDbConnection connection = new SqlConnection(_config.GetConnectionString("DefaultConnection"));
            string sql = @"SELECT CustCode, OrderNo, Proformas, Qty, ShippedQty, Price 
                           FROM VFOrderItemswithShippedQty 
                           WHERE CompItemCode = @ItemCode 
                           AND OrderNo IN(SELECT OrderNo FROM VUnshippedOrderList) 
                           AND Qty > ShippedQty 
                           AND OrderNo NOT IN('Stock-Order','Stock-OrderIss')";
            
            if (!string.IsNullOrEmpty(custCode) && custCode != "0")
            {
                sql += " AND CustCode = @CustCode";
            }
            
            sql += " ORDER BY OrderNo";

            return await connection.QueryAsync<ArticleUnshippedOrderModel>(sql, new { ItemCode = itemCode, CustCode = custCode });
        }

        public async Task<IEnumerable<ArticleCustomerPriceModel>> GetCustomerPricesAsync(string itemCode)
        {
            using IDbConnection connection = new SqlConnection(_config.GetConnectionString("DefaultConnection"));
            string sql = "SELECT * FROM FCustomerCatalog WHERE CompItemID = @ItemCode";
            return await connection.QueryAsync<ArticleCustomerPriceModel>(sql, new { ItemCode = itemCode });
        }

        public async Task<IEnumerable<ArticleQuotationModel>> GetQuotationsAsync(string itemCode)
        {
            using IDbConnection connection = new SqlConnection(_config.GetConnectionString("DefaultConnection"));
            string sql = "SELECT CustCode, QuotationNo, Qty, Price FROM VCustomerQuotations WHERE ItemCode = @ItemCode";
            return await connection.QueryAsync<ArticleQuotationModel>(sql, new { ItemCode = itemCode });
        }

        public async Task UpdateCustomerPriceAsync(long entryId, string columnName, double newValue, string userName)
        {
            // Validate columnName to prevent SQL injection
            var allowedColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase) 
            { 
                "FOB", "ExWorks", "CnISea", "CnIAir", "CIFSea", "CIFAir", "CnFSea", "CnFAir" 
            };

            if (!allowedColumns.Contains(columnName))
            {
                throw new ArgumentException("Invalid column name.");
            }

            using IDbConnection connection = new SqlConnection(_config.GetConnectionString("DefaultConnection"));
            connection.Open();
            using var transaction = connection.BeginTransaction();

            try
            {
                // Insert old values into history (FCustCatalRevs)
                string insertHistorySql = @"
                    INSERT INTO FCustCatalRevs(RefID, Description, FOB, CIFSea, CIFAir, CnFSea, CnFAir, FOBSialkot, CnISea, CnIAir, DT, UserName)
                    SELECT EntryID, Description, FOB, CIFSea, CIFAir, CnFSea, CnFAir, ExWorks, CnISea, CnIAir, GETDATE(), @UserName
                    FROM FCustomerCatalog 
                    WHERE EntryID = @EntryId";
                
                await connection.ExecuteAsync(insertHistorySql, new { EntryId = entryId, UserName = userName }, transaction);

                // Update new value in FCustomerCatalog
                string updateSql = $"UPDATE FCustomerCatalog SET {columnName} = @NewValue WHERE EntryID = @EntryId";
                await connection.ExecuteAsync(updateSql, new { NewValue = newValue, EntryId = entryId }, transaction);

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
