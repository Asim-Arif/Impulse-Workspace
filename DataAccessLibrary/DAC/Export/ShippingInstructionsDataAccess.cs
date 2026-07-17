using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Export
{
    public class ShippingInstructionsDataAccess : IShippingInstructionsDataAccess
    {
        private readonly string _connectionString;

        public ShippingInstructionsDataAccess(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
        }

        public async Task<List<CustomerLookupModel>> GetActiveCustomersAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT CustCode, CustCode AS Name FROM ForeignCustomers WHERE Active = 1 GROUP BY CustCode ORDER BY CustCode";
            var result = await db.QueryAsync<CustomerLookupModel>(sql);
            return result.ToList();
        }

        public async Task<List<string>> GetCountriesForCustomerAsync(string custCode)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT DISTINCT Country FROM ForeignCustomers WHERE CustCode = @CustCode ORDER BY Country";
            var result = await db.QueryAsync<string>(sql, new { CustCode = custCode });
            return result.ToList();
        }

        public async Task<List<string>> GetInvoicesWithInstructionsAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT CustomInvoice FROM VShippingInstructions1 WHERE CustCode = @CustCode AND Country = @Country ORDER BY CustomInvoice";
            var result = await db.QueryAsync<string>(sql, new { CustCode = custCode, Country = country });
            return result.ToList();
        }

        public async Task<List<string>> GetInvoicesWithoutInstructionsAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT CustomInvoice 
                FROM CustomInvoice 
                WHERE CustomInvoice NOT IN (SELECT CustomInvoice FROM ShippingInstructions) 
                  AND CustCode = @CustCode AND Country = @Country 
                ORDER BY CustomInvoice";
            var result = await db.QueryAsync<string>(sql, new { CustCode = custCode, Country = country });
            return result.ToList();
        }

        public async Task<ShippingInstructionsViewModel?> GetShippingInstructionsAsync(string customInvoice)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT CustomInvoice, GoodsDesc, GWeight, Measurement, MarksnNumbers, NonKind, Encloseres, SpecialInstructions 
                FROM ShippingInstructions 
                WHERE CustomInvoice = @CustomInvoice";
            return await db.QueryFirstOrDefaultAsync<ShippingInstructionsViewModel>(sql, new { CustomInvoice = customInvoice });
        }

        public async Task<(int TotalCartons, string InvoiceOf)> GetTotalCartonsAndInvoiceOfAsync(string customInvoice)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sqlCartons = "SELECT ISNULL(MAX(TotalCartons), 0) FROM VTotalCartons WHERE CustomInvoice = @CustomInvoice";
            const string sqlInvoiceOf = "SELECT ISNULL(MAX(InvoiceOf), '') FROM CustomInvoice WHERE CustomInvoice = @CustomInvoice";

            int totalCartons = await db.ExecuteScalarAsync<int>(sqlCartons, new { CustomInvoice = customInvoice });
            string invoiceOf = await db.ExecuteScalarAsync<string>(sqlInvoiceOf, new { CustomInvoice = customInvoice });

            return (totalCartons, invoiceOf);
        }

        public async Task<double> GetDefaultBoxSizeMeasurementAsync(int totalCartons)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT Dimension1 * Dimension2 * Dimension3 * @TotalCartons FROM BoxSize";
            var result = await db.QueryFirstOrDefaultAsync<double?>(sql, new { TotalCartons = totalCartons });
            return result ?? 0;
        }

        public async Task<bool> SaveShippingInstructionsAsync(ShippingInstructionsViewModel model)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            db.Open();
            using var transaction = db.BeginTransaction();
            try
            {
                // Delete existing first
                await db.ExecuteAsync(
                    "DELETE FROM ShippingInstructions WHERE CustomInvoice = @CustomInvoice",
                    new { CustomInvoice = model.CustomInvoice },
                    transaction: transaction);

                // Insert new
                const string sqlInsert = @"
                    INSERT INTO ShippingInstructions(CustomInvoice, GoodsDesc, GWeight, Measurement, MarksnNumbers, NonKind, Encloseres, SpecialInstructions) 
                    VALUES (@CustomInvoice, @GoodsDesc, @GWeight, @Measurement, @MarksnNumbers, @NonKind, @Encloseres, @SpecialInstructions)";
                
                await db.ExecuteAsync(sqlInsert, model, transaction: transaction);

                transaction.Commit();
                return true;
            }
            catch
            {
                transaction.Rollback();
                throw;
            }
        }
    }
}
