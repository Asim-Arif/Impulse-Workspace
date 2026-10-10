using Dapper;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.Extensions.Configuration;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Stock
{
    public class RMDataAccess : IRMDataAccess
    {
        private readonly string _connectionString;

        public RMDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<byte[]> GetMaterialPictureAsync(string materialId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = "SELECT Pic FROM RM WHERE RMID1 = @MaterialId";
                return await db.QueryFirstOrDefaultAsync<byte[]>(sql, new { MaterialId = materialId });
            }
        }

        public async Task<List<RMViewModel>> GetMaterialsAsync()
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT RMID1, '{' + RMID1 + '} ' + RMName AS RMName
                    FROM RM 
                    WHERE ISNULL(InActive, 0) = 0 AND LEFT(RMID1, 2) <> '32'
                    ORDER BY RMName";
                var results = await db.QueryAsync<RMViewModel>(sql);
                return results.AsList();
            }
        }

        public async Task<StockBreachSummaryDto> GetStockBreachSummaryAsync()
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT 
                        ISNULL(SUM(CASE WHEN ISNULL(QtyInStock, 0) <= ReorderPoint AND ISNULL(ReorderPoint, 0) > 0 AND (MinLevel IS NULL OR ISNULL(QtyInStock, 0) > MinLevel) THEN 1 ELSE 0 END), 0) AS BelowReorderCount,
                        ISNULL(SUM(CASE WHEN ISNULL(QtyInStock, 0) <= MinLevel AND ISNULL(MinLevel, 0) > 0 THEN 1 ELSE 0 END), 0) AS BelowMinCount,
                        ISNULL(SUM(CASE WHEN ISNULL(QtyInStock, 0) >= MaxLevel AND ISNULL(MaxLevel, 0) > 0 THEN 1 ELSE 0 END), 0) AS AboveMaxCount
                    FROM dbo.VRM
                    WHERE ISNULL(InActive, 0) = 0";
                var result = await db.QueryFirstOrDefaultAsync<StockBreachSummaryDto>(sql);
                return result ?? new StockBreachSummaryDto();
            }
        }

        public async Task<List<StockBreachItemDto>> GetStockBreachItemsAsync(string? breachType = null)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string whereClause = breachType?.ToLowerInvariant() switch
                {
                    "minstock" or "min" => "AND ISNULL(QtyInStock, 0) <= MinLevel AND ISNULL(MinLevel, 0) > 0",
                    "reorder" => "AND ISNULL(QtyInStock, 0) <= ReorderPoint AND ISNULL(ReorderPoint, 0) > 0 AND (MinLevel IS NULL OR ISNULL(QtyInStock, 0) > MinLevel)",
                    "maxstock" or "max" => "AND ISNULL(QtyInStock, 0) >= MaxLevel AND ISNULL(MaxLevel, 0) > 0",
                    _ => @"AND (
                            (ISNULL(QtyInStock, 0) <= MinLevel AND ISNULL(MinLevel, 0) > 0)
                            OR (ISNULL(QtyInStock, 0) <= ReorderPoint AND ISNULL(ReorderPoint, 0) > 0 AND (MinLevel IS NULL OR ISNULL(QtyInStock, 0) > MinLevel))
                            OR (ISNULL(QtyInStock, 0) >= MaxLevel AND ISNULL(MaxLevel, 0) > 0)
                          )"
                };

                string sql = $@"
                    SELECT 
                        RMID, RMID1, RMName, Unit, GroupName,
                        ISNULL(QtyInStock, 0) AS QtyInStock,
                        MinLevel, MaxLevel, ReorderPoint,
                        CASE 
                            WHEN ISNULL(QtyInStock, 0) <= MinLevel AND ISNULL(MinLevel, 0) > 0 THEN 'MinStock'
                            WHEN ISNULL(QtyInStock, 0) <= ReorderPoint AND ISNULL(ReorderPoint, 0) > 0 THEN 'Reorder'
                            WHEN ISNULL(QtyInStock, 0) >= MaxLevel AND ISNULL(MaxLevel, 0) > 0 THEN 'MaxStock'
                            ELSE 'Normal'
                        END AS BreachType
                    FROM dbo.VRM
                    WHERE ISNULL(InActive, 0) = 0 {whereClause}
                    ORDER BY 
                        CASE 
                            WHEN ISNULL(QtyInStock, 0) <= MinLevel AND ISNULL(MinLevel, 0) > 0 THEN 1
                            WHEN ISNULL(QtyInStock, 0) <= ReorderPoint AND ISNULL(ReorderPoint, 0) > 0 THEN 2
                            WHEN ISNULL(QtyInStock, 0) >= MaxLevel AND ISNULL(MaxLevel, 0) > 0 THEN 3
                            ELSE 4 
                        END, 
                        RMName";

                var results = await db.QueryAsync<StockBreachItemDto>(sql);
                return results.AsList();
            }
        }

        public async Task<StockBreachItemDto?> GetMaterialStockStatusByCodeAsync(string materialId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT TOP 1
                        RMID, RMID1, RMName, Unit, GroupName,
                        ISNULL(QtyInStock, 0) AS QtyInStock,
                        MinLevel, MaxLevel, ReorderPoint,
                        CASE 
                            WHEN ISNULL(QtyInStock, 0) <= MinLevel AND ISNULL(MinLevel, 0) > 0 THEN 'MinStock'
                            WHEN ISNULL(QtyInStock, 0) <= ReorderPoint AND ISNULL(ReorderPoint, 0) > 0 THEN 'Reorder'
                            WHEN ISNULL(QtyInStock, 0) >= MaxLevel AND ISNULL(MaxLevel, 0) > 0 THEN 'MaxStock'
                            ELSE 'Normal'
                        END AS BreachType
                    FROM dbo.VRM
                    WHERE RMID1 = @MaterialId";

                return await db.QueryFirstOrDefaultAsync<StockBreachItemDto>(sql, new { MaterialId = materialId });
            }
        }

        public async Task<StockBreachItemDto?> GetMaterialStockStatusByIdAsync(int rmId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT TOP 1
                        RMID, RMID1, RMName, Unit, GroupName,
                        ISNULL(QtyInStock, 0) AS QtyInStock,
                        MinLevel, MaxLevel, ReorderPoint,
                        CASE 
                            WHEN ISNULL(QtyInStock, 0) <= MinLevel AND ISNULL(MinLevel, 0) > 0 THEN 'MinStock'
                            WHEN ISNULL(QtyInStock, 0) <= ReorderPoint AND ISNULL(ReorderPoint, 0) > 0 THEN 'Reorder'
                            WHEN ISNULL(QtyInStock, 0) >= MaxLevel AND ISNULL(MaxLevel, 0) > 0 THEN 'MaxStock'
                            ELSE 'Normal'
                        END AS BreachType
                    FROM dbo.VRM
                    WHERE RMID = @RMID";

                return await db.QueryFirstOrDefaultAsync<StockBreachItemDto>(sql, new { RMID = rmId });
            }
        }
    }
}
