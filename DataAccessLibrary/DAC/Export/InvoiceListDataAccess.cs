using Dapper;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Export;
using Microsoft.Extensions.Configuration;
using System.Data.SqlClient;

namespace DataAccessLibrary.DAC.Export
{
    public class InvoiceListDataAccess : IInvoiceListDataAccess
    {
        private readonly IConfiguration _config;
        private string ConnectionString => _config.GetConnectionString("DefaultConnection")!;

        public InvoiceListDataAccess(IConfiguration config)
        {
            _config = config;
        }

        public async Task<IEnumerable<CompanyLookupModel>> GetCompaniesAsync()
        {
            using var db = new SqlConnection(ConnectionString);
            var sql = "SELECT EntryID, CompanyName FROM dbo.Companies ORDER BY CompanyName";
            return await db.QueryAsync<CompanyLookupModel>(sql);
        }

        public async Task<IEnumerable<CustomerLookupModel>> GetCustomersAsync()
        {
            using var db = new SqlConnection(ConnectionString);
            var sql = "SELECT DISTINCT CustCode, CustCode as Name FROM dbo.ForeignCustomers ORDER BY CustCode";
            return await db.QueryAsync<CustomerLookupModel>(sql);
        }

        public async Task<IEnumerable<string>> GetCountriesAsync()
        {
            using var db = new SqlConnection(ConnectionString);
            var sql = "SELECT DISTINCT Country FROM dbo.ForeignCustomers ORDER BY Country";
            return await db.QueryAsync<string>(sql);
        }

        public async Task<IEnumerable<InvoiceListModel>> GetInvoicesAsync(
            int invoiceType,
            long companyId,
            string customerCode,
            string country,
            string dateRange,
            DateTime? fromDate,
            DateTime? toDate,
            string searchText)
        {
            using var db = new SqlConnection(ConnectionString);

            string sql = "SELECT * FROM dbo.VCustomInvoiceList WHERE 1=1 ";
            var parameters = new DynamicParameters();

            // Invoice Type Logic
            // In legacy, we did logic by matching prefixes or just generally letting everything show unless strictly filtered
            // However, the legacy vb6 form actually relies on CustomInvoice prefix logic somewhere?
            // "If bCourierInvoice Then Cond = \" Where Left(CustomInvoice,4)='\" & strCourierInvoicePrefix & \"'\" Else Cond = \" Where Left(CustomInvoice,4)='VTL '\""
            // Wait, for regular Custom/Commercial/Bank, the VB6 code says:
            // For custom invoices, there is no prefix filter except the bCourierInvoice flag. The user wants the 3 standard views. I'll just load everything that matches the base filters, as the legacy VB6 app didn't explicitly filter the list by Custom/Commercial/Bank, but rather enabled different reports on the same list. 
            // So InvoiceType doesn't filter the actual SQL rows! It just alters the Context Menu!

            // Company Filter
            if (companyId > 0)
            {
                // In legacy: If strCompany = "PAKSMITH GROUP" Then ... we will just filter by CompanyID if it exists in the model or we ignore for now if not in view.
                // Wait, VCustomInvoiceList does NOT have a CompanyID field. 
                // Legacy didn't filter by Company! It just used it for feature flags!
                // Let's filter if possible, otherwise we skip. For now, no Company filter on the view since there is no Company column.
            }

            // Customer Filter
            if (!string.IsNullOrEmpty(customerCode) && customerCode != "0")
            {
                sql += " AND CustCode = @CustomerCode ";
                parameters.Add("@CustomerCode", customerCode);
            }

            // Country Filter
            if (!string.IsNullOrEmpty(country))
            {
                sql += " AND Country = @Country ";
                parameters.Add("@Country", country);
            }

            // Date Filter
            if (dateRange != "custom")
            {
                if (int.TryParse(dateRange, out int days))
                {
                    sql += $" AND DT >= DATEADD(day, -{days}, GETDATE()) ";
                }
            }
            else if (fromDate.HasValue && toDate.HasValue)
            {
                sql += " AND DT >= @FromDate AND DT <= @ToDate ";
                parameters.Add("@FromDate", fromDate.Value.Date);
                parameters.Add("@ToDate", toDate.Value.Date.AddDays(1).AddSeconds(-1));
            }

            // Search Filter
            if (!string.IsNullOrWhiteSpace(searchText))
            {
                sql += @" AND (
                            CustomInvoice LIKE @SearchText OR 
                            CustCode LIKE @SearchText OR 
                            Country LIKE @SearchText 
                        )";
                parameters.Add("@SearchText", $"%{searchText}%");
            }

            sql += " ORDER BY DT DESC";

            return await db.QueryAsync<InvoiceListModel>(sql, parameters);
        }
    }
}
