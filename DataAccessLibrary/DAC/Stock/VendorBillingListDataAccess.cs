using Dapper;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Stock
{
    public class VendorBillingListDataAccess : IVendorBillingListDataAccess
    {
        private readonly string _connectionString;

        public VendorBillingListDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<IEnumerable<VendorBillingListViewModel>> GetVendorBillingListAsync(DateTime fromDate, DateTime toDate, string vendorId, bool isSample)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT 
                        EntryID, 
                        VchrNo, 
                        BillNo, 
                        DT, 
                        AccTitle, 
                        ISNULL(TotalBillAmount, 0) AS TotalBillAmount, 
                        ISNULL(TotalDedAmount, 0) AS TotalDedAmount
                    FROM VVenderPostedBillsList
                    WHERE DT BETWEEN @FromDate AND @ToDate
                      AND SampleBilling = @IsSample";

                var parameters = new DynamicParameters();
                parameters.Add("FromDate", fromDate.Date);
                parameters.Add("ToDate", toDate.Date.AddDays(1).AddSeconds(-1));
                parameters.Add("IsSample", isSample ? 1 : 0);

                if (!string.IsNullOrEmpty(vendorId) && vendorId != "0")
                {
                    sql += " AND AccNo = @VendorId";
                    parameters.Add("VendorId", vendorId);
                }

                sql += " ORDER BY EntryID";

                return await db.QueryAsync<VendorBillingListViewModel>(sql, parameters);
            }
        }

        public async Task<IEnumerable<object>> GetVendorsAsync()
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = "SELECT AccNo, AccTitle FROM VVenders ORDER BY AccTitle";
                return await db.QueryAsync<dynamic>(sql);
            }
        }

        public async Task DeleteVendorBillAsync(int entryId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = "DELETE FROM VenderPostedBills WHERE EntryID = @EntryID";
                await db.ExecuteAsync(sql, new { EntryID = entryId });
            }
        }
    }
}
