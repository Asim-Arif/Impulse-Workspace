using Dapper;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Export
{
    public class ProformaListDataAccess : IProformaListDataAccess
    {
        private readonly IConfiguration _config;

        public ProformaListDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection") 
                                           ?? throw new InvalidOperationException("DefaultConnection not found.");

        public async Task<IEnumerable<ProformaListModel>> GetProformaListAsync(DateTime fromDate, DateTime toDate, string custCode, long companyId, string country)
        {
            using var db = new SqlConnection(ConnectionString);
            var sql = @"
                SELECT Curr, PInvoice, DT, CustCode, Country, ProformaAmt, Name as CompanyName, 
                       CompanyRefID, CompanyName, Orders, ChargesAmt, InternalOrders, AdvanceDescription, 
                       AdvanceAmount, DiscountAmount, Insurance_Freight_Amount, Insurance_Freight_Amount_II, DiscountPer
                FROM dbo.VFProformaList 
                WHERE DT BETWEEN @FromDate AND @ToDate";

            if (!string.IsNullOrEmpty(custCode) && custCode != "0")
            {
                sql += " AND CustCode = @CustCode";
            }
            if (companyId > 0)
            {
                sql += " AND CompanyRefID = @CompanyId";
            }
            if (!string.IsNullOrEmpty(country))
            {
                sql += " AND Country = @Country";
            }

            sql += " ORDER BY DT DESC, PInvoice";

            return await db.QueryAsync<ProformaListModel>(sql, new { FromDate = fromDate, ToDate = toDate, CustCode = custCode, CompanyId = companyId, Country = country });
        }

        public async Task<IEnumerable<CustomerLookupModel>> GetCustomersAsync()
        {
            using var db = new SqlConnection(ConnectionString);
            var sql = "SELECT CustCode, Name FROM dbo.ForeignCustomers WHERE CustCode <> 'Hico' ORDER BY Name";
            return await db.QueryAsync<CustomerLookupModel>(sql);
        }

        public async Task<IEnumerable<CompanyLookupModel>> GetCompaniesAsync()
        {
            using var db = new SqlConnection(ConnectionString);
            var sql = "SELECT EntryID, CompanyName FROM dbo.Companies ORDER BY CompanyName";
            return await db.QueryAsync<CompanyLookupModel>(sql);
        }

        public async Task<IEnumerable<string>> GetCountriesAsync(string custCode)
        {
            using var db = new SqlConnection(ConnectionString);
            var sql = "SELECT DISTINCT Country FROM dbo.ForeignCustomers WHERE 1=1";
            if (!string.IsNullOrEmpty(custCode) && custCode != "0")
            {
                sql += " AND CustCode = @CustCode";
            }
            return await db.QueryAsync<string>(sql, new { CustCode = custCode });
        }

        public async Task<bool> DeleteProformaAsync(string pInvoice)
        {
            using var db = new SqlConnection(ConnectionString);
            var checkSql = @"
                SELECT COUNT(*) 
                FROM FProformaOrders 
                INNER JOIN CustomInvoiceItems ON FProformaOrders.EntryID=CustomInvoiceItems.RefID 
                WHERE FProformaOrders.PInvoice = @PInvoice";
            
            var count = await db.ExecuteScalarAsync<int>(checkSql, new { PInvoice = pInvoice });
            if (count > 0)
            {
                throw new Exception("One or more invoices are using the selected proforma. Cannot Delete.");
            }

            var deleteSql = "DELETE FROM FPInvoice WHERE PInvoice = @PInvoice";
            var rows = await db.ExecuteAsync(deleteSql, new { PInvoice = pInvoice });
            return rows > 0;
        }
    }
}
