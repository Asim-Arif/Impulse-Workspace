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
    }
}
