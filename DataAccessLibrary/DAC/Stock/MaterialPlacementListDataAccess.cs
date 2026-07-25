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
    public class MaterialPlacementListDataAccess : IMaterialPlacementListDataAccess
    {
        private readonly string _connectionString;

        public MaterialPlacementListDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<IEnumerable<PlacedMaterialItemViewModel>> GetPlacedMaterialsListAsync(DateTime fromDate, DateTime toDate, string vendorId, string materialId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT 
                        EntryID, 
                        RcvDate, 
                        AccTitle as VendorName, 
                        PORefNo, 
                        RMName as MaterialName, 
                        StoreName, 
                        RackNo, 
                        ShelfNo, 
                        QtyPlaced, 
                        Unit, 
                        RMID1 
                    FROM VMaterialLocationWiseStatus
                    WHERE RcvDate BETWEEN @FromDate AND @ToDate";

                var parameters = new DynamicParameters();
                parameters.Add("FromDate", fromDate.Date);
                parameters.Add("ToDate", toDate.Date.AddDays(1).AddSeconds(-1)); // End of day

                if (!string.IsNullOrEmpty(vendorId) && vendorId != "0")
                {
                    sql += " AND AccNo = @VendorId";
                    parameters.Add("VendorId", vendorId);
                }

                if (!string.IsNullOrEmpty(materialId) && materialId != "0")
                {
                    sql += " AND RMID1 = @MaterialId";
                    parameters.Add("MaterialId", materialId);
                }

                sql += " ORDER BY EntryID";

                return await db.QueryAsync<PlacedMaterialItemViewModel>(sql, parameters);
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

        public async Task<IEnumerable<object>> GetMaterialsAsync()
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = "SELECT RMID1, RMID1 + ' {'+RMName+'}' as MaterialName FROM VRM ORDER BY RMName";
                return await db.QueryAsync<dynamic>(sql);
            }
        }
    }
}
