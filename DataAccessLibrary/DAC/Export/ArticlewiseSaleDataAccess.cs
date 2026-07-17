using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Export;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Export
{
    public class ArticlewiseSaleDataAccess : IArticlewiseSaleDataAccess
    {
        private readonly string _connectionString;

        public ArticlewiseSaleDataAccess(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
        }

        public async Task<List<ItemLookupModel>> GetArticlesAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT ItemID, ItemID + ' ' + ItemName + ' ' + ISNULL(ItemSize, '') AS ItemName FROM Items ORDER BY ItemID";
            var result = await db.QueryAsync<ItemLookupModel>(sql);
            return result.ToList();
        }

        public async Task<List<CustomerLookupModel>> GetActiveCustomersAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT DISTINCT CustCode, CustCode AS Name FROM ForeignCustomers ORDER BY CustCode";
            var result = await db.QueryAsync<CustomerLookupModel>(sql);
            return result.ToList();
        }

        public async Task<List<string>> GetCountriesAsync(string custCode)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql;
            
            if (string.IsNullOrEmpty(custCode) || custCode == "0")
            {
                sql = "SELECT DISTINCT Country FROM ForeignCustomers WHERE ISNULL(Country, '') <> '' ORDER BY Country";
                var result = await db.QueryAsync<string>(sql);
                return result.ToList();
            }
            else
            {
                sql = "SELECT DISTINCT Country FROM ForeignCustomers WHERE CustCode = @CustCode AND ISNULL(Country, '') <> '' ORDER BY Country";
                var result = await db.QueryAsync<string>(sql, new { CustCode = custCode });
                return result.ToList();
            }
        }

        public async Task<List<ArticlewiseSaleModel>> GetArticlewiseSalesDataAsync(string itemId, string custCode, string country, bool useDateRange, DateTime dtFrom, DateTime dtTo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            
            var conditions = new List<string>();
            var parameters = new DynamicParameters();

            if (!string.IsNullOrEmpty(itemId) && itemId != "0")
            {
                conditions.Add("ItemID = @ItemID");
                parameters.Add("ItemID", itemId);
            }

            if (!string.IsNullOrEmpty(custCode) && custCode != "0")
            {
                conditions.Add("CustCode = @CustCode");
                parameters.Add("CustCode", custCode);
            }

            if (!string.IsNullOrEmpty(country) && country != "0")
            {
                conditions.Add("Country = @Country");
                parameters.Add("Country", country);
            }

            if (useDateRange)
            {
                conditions.Add("DT BETWEEN @DtFrom AND @DtTo");
                parameters.Add("DtFrom", dtFrom.Date);
                parameters.Add("DtTo", dtTo.Date.AddDays(1).AddTicks(-1));
            }

            string sql = "SELECT CustCode, Country, CustomInvoice, DT, Qty, ItemID, ItemName FROM VArticlewiseSales";
            if (conditions.Count > 0)
            {
                sql += " WHERE " + string.Join(" AND ", conditions);
            }
            sql += " ORDER BY DT";

            var result = await db.QueryAsync<ArticlewiseSaleModel>(sql, parameters);
            return result.ToList();
        }
    }
}
