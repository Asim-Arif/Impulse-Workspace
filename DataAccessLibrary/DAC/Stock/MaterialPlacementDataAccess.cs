using Dapper;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Stock
{
    public class MaterialPlacementDataAccess : IMaterialPlacementDataAccess
    {
        private readonly string _connectionString;

        public MaterialPlacementDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<List<MaterialPlacementListViewModel>> GetPendingPlacementsAsync(string vendorId, string materialId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var queryBuilder = new System.Text.StringBuilder(
                    "SELECT EntryID, RcvID, MaterialID, RMName, AccTitle AS VendorName, " +
                    "ISNULL(QtyPassed, 0) AS QtyPassed, ISNULL(QtyPlaced, 0) AS QtyPlaced, " +
                    "(ISNULL(QtyPassed, 0) - ISNULL(QtyPlaced, 0)) AS Balance, Unit, " +
                    "CCItem AS IsCCItem, RISD_EntryID, RcvDate " +
                    "FROM VMaterialsForPlacement " +
                    "WHERE ROUND(QtyPassed, 2) > ROUND(QtyPlaced, 2)");

                var parameters = new DynamicParameters();

                if (!string.IsNullOrEmpty(materialId) && materialId != "0")
                {
                    queryBuilder.Append(" AND RMID1 = @MaterialID");
                    parameters.Add("MaterialID", materialId);
                }

                if (!string.IsNullOrEmpty(vendorId) && vendorId != "0")
                {
                    queryBuilder.Append(" AND VendID = @VendID");
                    parameters.Add("VendID", vendorId);
                }

                queryBuilder.Append(" ORDER BY RcvDate, EntryID");

                var results = await db.QueryAsync<MaterialPlacementListViewModel>(queryBuilder.ToString(), parameters);
                return results.ToList();
            }
        }

        public async Task<List<StoreViewModel>> GetStoresForUserAsync(int userId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT EntryID, StoreName 
                    FROM Stores 
                    WHERE EntryID IN (SELECT Store_RefID FROM Users_Stores WHERE UserID = @UserID)
                    ORDER BY StoreName";
                var results = await db.QueryAsync<StoreViewModel>(sql, new { UserID = userId });
                var list = results.ToList();
                if (!list.Any())
                {
                    // Fallback for testing if user has no mapped stores
                    string fallbackSql = "SELECT EntryID, StoreName FROM Stores ORDER BY StoreName";
                    var fallbackResults = await db.QueryAsync<StoreViewModel>(fallbackSql);
                    return fallbackResults.ToList();
                }
                return list;
            }
        }

        public async Task<List<RackViewModel>> GetRacksAsync(int storeId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = "SELECT EntryID, RackNo FROM StoreRacks WHERE Store_RefID = @StoreId ORDER BY RackNo";
                var results = await db.QueryAsync<RackViewModel>(sql, new { StoreId = storeId });
                return results.ToList();
            }
        }

        public async Task<List<ShelfViewModel>> GetShelvesAsync(int rackId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = "SELECT EntryID, ShelfNo FROM StoreShelfs WHERE Rack_RefID = @RackId ORDER BY ShelfNo";
                var results = await db.QueryAsync<ShelfViewModel>(sql, new { RackId = rackId });
                return results.ToList();
            }
        }

        public async Task SavePlacementAsync(PlacementSaveRequest request, string userName, string machineName)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                db.Open();
                using (var transaction = db.BeginTransaction())
                {
                    try
                    {
                        string sqlMaster = @"
                            INSERT INTO MaterialLocationwiseStatus
                            (Rcvd_RefID, Shelf_RefID, QtyPlaced, UserName, MachineName, BatchNo, LotNo) 
                            VALUES
                            (@Rcvd_RefID, @Shelf_RefID, @QtyPlaced, @UserName, @MachineName, @BatchNo, @LotNo)";

                        await db.ExecuteAsync(sqlMaster, new
                        {
                            request.Rcvd_RefID,
                            request.Shelf_RefID,
                            request.QtyPlaced,
                            UserName = userName,
                            MachineName = machineName,
                            request.BatchNo,
                            request.LotNo
                        }, transaction);

                        if (request.IsCCItem && request.RISD_RefID > 0)
                        {
                            string sqlDetail = @"
                                INSERT INTO RcvItemsSimpleDetail_Placement
                                (RISD_RefID, Shelf_RefID, RcvdQty, Remarks) 
                                VALUES
                                (@RISD_RefID, @Shelf_RefID, @QtyPlaced, '')";
                                
                            await db.ExecuteAsync(sqlDetail, new
                            {
                                request.RISD_RefID,
                                request.Shelf_RefID,
                                request.QtyPlaced
                            }, transaction);
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
    }
}
