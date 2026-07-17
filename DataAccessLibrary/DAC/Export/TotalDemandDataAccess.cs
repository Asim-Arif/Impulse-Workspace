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
    public class TotalDemandDataAccess : ITotalDemandDataAccess
    {
        private readonly string _connectionString;

        public TotalDemandDataAccess(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
        }

        public async Task<List<CustomerLookupModel>> GetActiveCustomersAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT CustCode, CustCode AS Name FROM ForeignCustomers WHERE CustCode <> 'HICO' GROUP BY CustCode ORDER BY CustCode";
            var result = await db.QueryAsync<CustomerLookupModel>(sql);
            return result.ToList();
        }

        public async Task<List<TotalDemandModel>> GetTotalDemandDataAsync(string custCode, DateTime dtFrom, DateTime dtTo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            
            // Build base SQL query
            string sql = @"
                SELECT CustCode, Country, OrderNo, DT, OrderAmt, OrderType 
                FROM VFOrderList
                WHERE DT BETWEEN @DtFrom AND @DtTo
                  AND OrderNo NOT IN (SELECT OrderNo FROM FCustomerFinalOrders WHERE Cancelled IN (1, 2))
                  AND OrderType = 'Customer Order'
                  AND CustCode NOT IN ('HMP')";

            // Conditional customer filter
            if (!string.IsNullOrEmpty(custCode) && custCode != "0")
            {
                sql += " AND CustCode = @CustCode";
            }

            sql += " ORDER BY DT, OrderNo";

            var result = await db.QueryAsync<TotalDemandModel>(sql, new { 
                CustCode = custCode, 
                DtFrom = dtFrom.Date, 
                DtTo = dtTo.Date.AddDays(1).AddTicks(-1) // End of day
            });

            return result.ToList();
        }
    }
}
