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
    public class TotalExportDataAccess : ITotalExportDataAccess
    {
        private readonly string _connectionString;

        public TotalExportDataAccess(IConfiguration configuration)
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

        public async Task<List<TotalExportModel>> GetTotalExportDataAsync(string custCode, DateTime dtFrom, DateTime dtTo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            
            // Build base SQL query
            string sql = @"
                SELECT CustCode, Country, CustomInvoice, DT, TotalAmt, TotalCustomAmt, ExchRate, GatePassNo 
                FROM VTotalExport
                WHERE DT BETWEEN @DtFrom AND @DtTo
                  AND ISNULL(GatePassNo, '') <> ''
                  AND CustCode NOT IN ('HMP')";

            // Conditional customer filter
            if (!string.IsNullOrEmpty(custCode) && custCode != "0")
            {
                sql += " AND CustCode = @CustCode";
            }

            sql += " ORDER BY DT, CustomInvoice";

            var result = await db.QueryAsync<TotalExportModel>(sql, new { 
                CustCode = custCode, 
                DtFrom = dtFrom.Date, 
                DtTo = dtTo.Date.AddDays(1).AddTicks(-1) // End of day
            });

            return result.ToList();
        }
    }
}
