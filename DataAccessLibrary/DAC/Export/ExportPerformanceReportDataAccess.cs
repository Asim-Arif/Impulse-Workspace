using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Export
{
    public class ExportPerformanceReportDataAccess : IExportPerformanceReportDataAccess
    {
        private readonly string _connectionString;

        public ExportPerformanceReportDataAccess(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new System.InvalidOperationException("Connection string 'DefaultConnection' not found.");
        }

        public async Task<List<CustomerLookupModel>> GetCustomersAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT DISTINCT CustCode, CustCode AS Name FROM ForeignCustomers ORDER BY CustCode";
            return (await db.QueryAsync<CustomerLookupModel>(sql)).ToList();
        }

        public async Task<List<ExportItemGroupLookupModel>> GetItemGroupsAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT ID, Description FROM ItemGroups ORDER BY Description";
            return (await db.QueryAsync<ExportItemGroupLookupModel>(sql)).ToList();
        }

        public async Task<List<ExportItemCategoryLookupModel>> GetItemCategoriesAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT CatID, Description FROM ItemCatagories ORDER BY Description";
            return (await db.QueryAsync<ExportItemCategoryLookupModel>(sql)).ToList();
        }

        public async Task<List<ExportMainGroupLookupModel>> GetMainGroupsAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "SELECT MainGroupID, MainGroupName FROM ItemGroupsMain ORDER BY MainGroupName";
            return (await db.QueryAsync<ExportMainGroupLookupModel>(sql)).ToList();
        }
    }
}
