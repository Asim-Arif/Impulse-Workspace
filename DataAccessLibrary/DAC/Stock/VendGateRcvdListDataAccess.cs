using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using Dapper;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using DataAccessLibrary.Models.ViewModels.Stock;
using DataAccessLibrary.Interface.Stock;

namespace DataAccessLibrary.DAC.Stock
{
    public class VendGateRcvdListDataAccess : IVendGateRcvdListDataAccess
    {
        private readonly IConfiguration _config;

        public VendGateRcvdListDataAccess(IConfiguration config)
        {
            _config = config;
        }

        public async Task<List<object>> GetVendorsAsync()
        {
            using IDbConnection connection = new SqlConnection(_config.GetConnectionString("DefaultConnection"));
            string sql = "SELECT AccNo, AccTitle FROM VVenders ORDER BY AccTitle";
            var result = await connection.QueryAsync(sql);
            return result.AsList();
        }

        public async Task<List<object>> GetRMGroupsAsync()
        {
            using IDbConnection connection = new SqlConnection(_config.GetConnectionString("DefaultConnection"));
            string sql = "SELECT ID, Description FROM RMGroups ORDER BY Description";
            var result = await connection.QueryAsync(sql);
            return result.AsList();
        }

        public async Task<List<VendGateRcvdListViewModel>> GetReceivingListAsync(DateTime fromDate, DateTime toDate, string vendorId, string poNo, string groupId, bool sampleOrder)
        {
            using IDbConnection connection = new SqlConnection(_config.GetConnectionString("DefaultConnection"));
            
            var parameters = new DynamicParameters();
            parameters.Add("@FromDate", fromDate);
            parameters.Add("@ToDate", toDate);
            parameters.Add("@SampleOrder", sampleOrder ? 1 : 0);

            string sql = @"
                SELECT RcvID, PORefNo, RcvDate, VenderName, ISNULL(TotalQty, 0) AS TotalQty, ISNULL(TotalQtyRejected, 0) AS TotalQtyRejected 
                FROM VVenderRcvs 
                WHERE (RcvDate BETWEEN @FromDate AND @ToDate) 
                  AND SampleOrder = @SampleOrder";

            if (!string.IsNullOrEmpty(vendorId) && vendorId != "0")
            {
                sql += " AND VendID = @VendorId";
                parameters.Add("@VendorId", vendorId);
            }

            if (!string.IsNullOrEmpty(poNo))
            {
                sql += " AND PORefNo LIKE '%' + @PONo + '%'";
                parameters.Add("@PONo", poNo);
            }

            if (!string.IsNullOrEmpty(groupId) && groupId != "0")
            {
                sql += " AND RcvID IN(SELECT RcvID FROM VVendRcvdDetailPO WHERE GroupID = @GroupID)";
                parameters.Add("@GroupID", groupId);
            }

            sql += " ORDER BY RcvDate ASC";

            var result = await connection.QueryAsync<VendGateRcvdListViewModel>(sql, parameters);
            
            var list = result.AsList();
            for (int i = 0; i < list.Count; i++)
            {
                list[i].SNo = i + 1;
            }
            
            return list;
        }

        public async Task DeleteReceivingAsync(string rcvId)
        {
            using IDbConnection connection = new SqlConnection(_config.GetConnectionString("DefaultConnection"));
            connection.Open();
            using var transaction = connection.BeginTransaction();

            try
            {
                // Lock placements and check if any material from this receiving has been issued
                string checkSql = @"
                    SELECT COUNT(*)
                    FROM MaterialLocationwiseStatus mls WITH (UPDLOCK, HOLDLOCK)
                    INNER JOIN VendRcvdDetailPO vrdp ON mls.Rcvd_RefID = vrdp.EntryID
                    WHERE vrdp.RcvID = @RcvID
                      AND ISNULL(mls.QtyIssued, 0) > 0";

                int issuedCount = await connection.ExecuteScalarAsync<int>(checkSql, new { RcvID = rcvId }, transaction);
                if (issuedCount > 0)
                {
                    throw new Exception("Can't Delete, Material Issued.");
                }

                // Reverse RM.QtyInStock for checked receivings before cascading delete
                string reverseRmSql = @"
                    UPDATE rm
                    SET rm.QtyInStock = ISNULL(rm.QtyInStock, 0) - Agg.NetQty
                    FROM RM rm
                    INNER JOIN (
                        SELECT MaterialID,
                               SUM(ISNULL(QtyRcvd, 0) - ISNULL(QtyReturned, 0) + ISNULL(ExtraQtyRcvd, 0)) AS NetQty
                        FROM VendRcvdDetailPO
                        WHERE RcvID = @RcvID AND ISNULL(Checked, 0) = 1
                        GROUP BY MaterialID
                    ) Agg ON rm.RMID1 = Agg.MaterialID";
                await connection.ExecuteAsync(reverseRmSql, new { RcvID = rcvId }, transaction);

                // Delete header (cascades VendRcvdDetailPO and unissued MaterialLocationwiseStatus)
                string deleteSql = "DELETE FROM VendRcvd WHERE RcvID = @RcvID";
                await connection.ExecuteAsync(deleteSql, new { RcvID = rcvId }, transaction);

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
