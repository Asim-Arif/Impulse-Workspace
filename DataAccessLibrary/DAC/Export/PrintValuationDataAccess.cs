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
    public class PrintValuationDataAccess : IPrintValuationDataAccess
    {
        private readonly string _connectionString;

        public PrintValuationDataAccess(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
        }

        public async Task<List<CustomerLookupModel>> GetCustomersAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT CustCode, CustCode AS Name FROM ForeignCustomers WHERE CustCode <> 'HICO' GROUP BY CustCode ORDER BY CustCode";
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

        public async Task<List<string>> GetInvoicesAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT DISTINCT CustomInvoice FROM VrptCustomInvoice WHERE CustCode = @CustCode AND Country = @Country ORDER BY CustomInvoice";
            var result = await db.QueryAsync<string>(sql, new { CustCode = custCode, Country = country });
            return result.ToList();
        }
    }
}
