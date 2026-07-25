using Dapper;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.Extensions.Configuration;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Stock
{
    public class ChangeBatchNoLotDataAccess : IChangeBatchNoLotDataAccess
    {
        private readonly string _connectionString;

        public ChangeBatchNoLotDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<List<ItemViewModel>> GetItemsAsync()
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT ItemID, '{' + ItemID + '} ' + ItemName AS ItemName
                    FROM Items 
                    WHERE ISNULL(InActive, 0) = 0
                    ORDER BY ItemID";
                var results = await db.QueryAsync<ItemViewModel>(sql);
                return results.AsList();
            }
        }

        public async Task<List<LotListViewModel>> GetLotsByMaterialAsync(string materialId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT EntryID, LotNo, ItemID, Batch_No, Mill_Certificate_No, Lot_Remarks 
                    FROM Lots_List 
                    WHERE ItemID = @MaterialId 
                    ORDER BY EntryID";
                var results = await db.QueryAsync<LotListViewModel>(sql, new { MaterialId = materialId });
                return results.AsList();
            }
        }

        public async Task UpdateLotsAsync(List<LotListViewModel> lots, string userName, string machineName)
        {
            var modifiedLots = lots.Where(l => l.IsModified).ToList();
            if (!modifiedLots.Any())
                return;

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                db.Open();
                using (var transaction = db.BeginTransaction())
                {
                    try
                    {
                        string updateSql = @"
                            UPDATE Lots_List 
                            SET Batch_No = @Batch_No, 
                                Mill_Certificate_No = @Mill_Certificate_No, 
                                Lot_Remarks = @Lot_Remarks 
                            WHERE EntryID = @EntryID";

                        await db.ExecuteAsync(updateSql, modifiedLots, transaction);

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
