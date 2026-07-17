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
    public class PackingLabelsDataAccess : IPackingLabelsDataAccess
    {
        private readonly string _connectionString;

        public PackingLabelsDataAccess(IConfiguration configuration)
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

        public async Task<CustomerLabelSettingsModel?> GetLabelSettingsAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT Curr, OuterPackingLabel, InnerPackingLabel FROM ForeignCustomers WHERE CustCode = @CustCode AND Country = @Country";
            return await db.QueryFirstOrDefaultAsync<CustomerLabelSettingsModel>(sql, new { CustCode = custCode, Country = country });
        }

        public async Task<List<string>> GetCustomInvoicesAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT CustomInvoice FROM VrptCustomInvoice WHERE CustCode = @CustCode AND Country = @Country ORDER BY CustomInvoice";
            var result = await db.QueryAsync<string>(sql, new { CustCode = custCode, Country = country });
            return result.ToList();
        }

        public async Task<int> GetTotalCartonsAsync(string customInvoice)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT ISNULL(MAX(TotalCartons), 0) FROM VTotalCartons WHERE CustomInvoice = @CustomInvoice";
            return await db.ExecuteScalarAsync<int>(sql, new { CustomInvoice = customInvoice });
        }

        public async Task GeneratePrintPackingLabelsAsync(string customInvoice)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            db.Open();
            using var transaction = db.BeginTransaction();
            try
            {
                // 1. Delete all existing records
                await db.ExecuteAsync("DELETE FROM PrintPackingLabels", transaction: transaction);

                // 2. Fetch carton ranges from CustomPList
                const string sqlPList = @"
                    SELECT ID, CartonFrom, CartonTo 
                    FROM CustomPList 
                    WHERE CustomInvoice = @CustomInvoice 
                    ORDER BY CartonFrom, ID";
                
                var items = (await db.QueryAsync<(int ID, int CartonFrom, int CartonTo)>(
                    sqlPList, 
                    new { CustomInvoice = customInvoice }, 
                    transaction: transaction)).ToList();

                // 3. Populate PrintPackingLabels
                const string sqlInsert = "INSERT INTO PrintPackingLabels (RefID, CartonNo) VALUES (@RefID, @CartonNo)";
                var insertParams = new List<object>();

                foreach (var item in items)
                {
                    for (int i = item.CartonFrom; i <= item.CartonTo; i++)
                    {
                        insertParams.Add(new { RefID = item.ID, CartonNo = i });
                    }
                }

                if (insertParams.Any())
                {
                    await db.ExecuteAsync(sqlInsert, insertParams, transaction: transaction);
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
