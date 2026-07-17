using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Company;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Export
{
    public class PrintInnerLabelsDataAccess : IPrintInnerLabelsDataAccess
    {
        private readonly string _connectionString;

        public PrintInnerLabelsDataAccess(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
        }

        public async Task<List<ItemLookupModel>> GetActiveItemsAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT ItemID, '{' + ItemID + '} ' + ItemName AS ItemName 
                FROM Items 
                WHERE InActive = 0 
                ORDER BY ItemID";
            var result = await db.QueryAsync<ItemLookupModel>(sql);
            return result.ToList();
        }

        public async Task<List<string>> GetOrderNumbersForItemAsync(string itemId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT DISTINCT OrderNo 
                FROM FOrderItems 
                WHERE CompItemCode = @CompItemCode 
                ORDER BY OrderNo";
            var result = await db.QueryAsync<string>(sql, new { CompItemCode = itemId });
            return result.ToList();
        }
    }
}
