using Dapper;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Production
{
    public class MakerListDataAccess : IMakerListDataAccess
    {
        private readonly IConfiguration _config;

        public MakerListDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection configuration string is missing.");

        public async Task<List<MakerDto>> GetMakersListAsync(bool includeInactive)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT 
                            VendID, VendID1, VenderName, CompanyName, Maker_Second_Name,
                            Phone1, Phone2, Fax1 AS Phone3, MakerType, Active, AccNo
                           FROM VMakers 
                           WHERE (@IncludeInactive = 1 OR ISNULL(Active, 1) = 1)
                           ORDER BY VenderName";

            var list = (await db.QueryAsync<MakerDto>(sql, new { IncludeInactive = includeInactive ? 1 : 0 })).ToList();
            return list;
        }

        public async Task ToggleMakerActiveStatusAsync(long vendId, bool newActiveState)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"UPDATE Makers SET Active = @NewActiveState WHERE VendID = @VendId";
            await db.ExecuteAsync(sql, new { VendId = vendId, NewActiveState = newActiveState ? 1 : 0 });
        }
    }
}
