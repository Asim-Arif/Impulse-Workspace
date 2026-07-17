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
    public class CommercialCoveringDataAccess : ICommercialCoveringDataAccess
    {
        private readonly string _connectionString;

        public CommercialCoveringDataAccess(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
        }

        public async Task<List<CustomerLookupModel>> GetActiveCustomersAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT CustCode, CustCode AS Name FROM ForeignCustomers WHERE CustCode <> 'Hico' GROUP BY CustCode ORDER BY CustCode";
            var result = await db.QueryAsync<CustomerLookupModel>(sql);
            return result.ToList();
        }

        public async Task<List<CustomerCountryModel>> GetCountriesForCustomerAsync(string custCode)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT Country, Curr FROM ForeignCustomers WHERE CustCode = @CustCode ORDER BY Country";
            var result = await db.QueryAsync<CustomerCountryModel>(sql, new { CustCode = custCode });
            return result.ToList();
        }

        public async Task<List<string>> GetInvoicesWithLettersAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT DISTINCT CustomInvoice FROM VCommercialCoveringLetter WHERE CustCode = @CustCode AND Country = @Country ORDER BY CustomInvoice";
            var result = await db.QueryAsync<string>(sql, new { CustCode = custCode, Country = country });
            return result.ToList();
        }

        public async Task<List<string>> GetInvoicesWithoutLettersAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT CustomInvoice 
                FROM CustomInvoice 
                WHERE CustomInvoice NOT IN (SELECT CustomInvoice FROM CommercialCoveringLetter) 
                  AND CustCode = @CustCode AND Country = @Country 
                ORDER BY CustomInvoice";
            var result = await db.QueryAsync<string>(sql, new { CustCode = custCode, Country = country });
            return result.ToList();
        }

        public async Task<List<CoveringLetterItemViewModel>> GetCoveringLetterItemsAsync(string customInvoice)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT EntryText, EntryDetail FROM VCommercialCoveringLetter WHERE CustomInvoice = @CustomInvoice ORDER BY EntryID";
            var result = await db.QueryAsync<CoveringLetterItemViewModel>(sql, new { CustomInvoice = customInvoice });
            return result.ToList();
        }

        public async Task<bool> SaveCoveringLetterItemsAsync(string customInvoice, List<CoveringLetterItemViewModel> items)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            db.Open();
            using var transaction = db.BeginTransaction();
            try
            {
                // Delete existing items
                await db.ExecuteAsync(
                    "DELETE FROM CommercialCoveringLetter WHERE CustomInvoice = @CustomInvoice",
                    new { CustomInvoice = customInvoice },
                    transaction: transaction);

                // Insert new items
                const string sqlInsert = @"
                    INSERT INTO CommercialCoveringLetter(CustomInvoice, EntryText, EntryDetail) 
                    VALUES (@CustomInvoice, @EntryText, @EntryDetail)";

                foreach (var item in items)
                {
                    // Skip empty rows if any
                    if (string.IsNullOrWhiteSpace(item.EntryText))
                        continue;

                    await db.ExecuteAsync(sqlInsert, new { 
                        CustomInvoice = customInvoice, 
                        EntryText = item.EntryText, 
                        EntryDetail = item.EntryDetail 
                    }, transaction: transaction);
                }

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
