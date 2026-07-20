using Dapper;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels.Export;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;

namespace DataAccessLibrary.DAC.Accounts
{
    public class CustomInvoiceAuthDataAccess : ICustomInvoiceAuthDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<CustomInvoiceAuthDataAccess> _logger;
        private readonly IVouchersDataAccess _vouchersDataAccess;

        public CustomInvoiceAuthDataAccess(IConfiguration configuration, ILogger<CustomInvoiceAuthDataAccess> logger, IVouchersDataAccess vouchersDataAccess)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                            ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
            _vouchersDataAccess = vouchersDataAccess;
        }

        // Legacy: cmbCust.AddVals(con, "Distinct CustCode", "ForeignCustomers", "CustCode")
        public async Task<IEnumerable<CustomerLookupModel>> GetCustomersAsync()
        {
            using var db = new SqlConnection(_connectionString);
            var sql = "SELECT DISTINCT CustCode, CustCode as Name FROM dbo.ForeignCustomers ORDER BY CustCode";
            return await db.QueryAsync<CustomerLookupModel>(sql);
        }

        // Legacy cmbCust_matched: Select Distinct Country From ForeignCustomers [Where CustCode=...]
        public async Task<IEnumerable<string>> GetCountriesAsync(string? customerCode)
        {
            using var db = new SqlConnection(_connectionString);
            var sql = "SELECT DISTINCT Country FROM dbo.ForeignCustomers";
            var parameters = new DynamicParameters();
            if (!string.IsNullOrEmpty(customerCode) && customerCode != "0")
            {
                sql += " WHERE CustCode = @CustomerCode";
                parameters.Add("@CustomerCode", customerCode);
            }
            sql += " ORDER BY Country";
            return await db.QueryAsync<string>(sql, parameters);
        }

        // Legacy Form_Load: AddToCombo CmbAccTitle, "AccTitle", "Accounts", " WHERE SubAccOf='31-001' OR AccNo='33-001-12001'", , "AccNo"
        public async Task<IEnumerable<ChartOfAccountsModel>> GetSalesAccountsAsync()
        {
            using var db = new SqlConnection(_connectionString);
            var sql = "SELECT AccNo, AccTitle FROM dbo.Accounts WHERE SubAccOf='31-001' OR AccNo='33-001-12001'";
            return await db.QueryAsync<ChartOfAccountsModel>(sql);
        }

        // Legacy RefreshLV: Select * From VCustomInvoiceAuth WHERE DT BETWEEN DtFrom AND DtTo [AND CustCode=...] [AND Country=...] Order By DT
        public async Task<IEnumerable<CustomInvoiceAuthItemModel>> GetPendingInvoicesAsync(string? customerCode, string? country, DateTime fromDate, DateTime toDate)
        {
            using var db = new SqlConnection(_connectionString);
            var sql = @"SELECT Curr, CustomInvoice, DT, CustCode, Country, TotalAmt, PackingInfo, ISNULL(AccNo, '') AS AccNo, TotalCustomAmt
                        FROM dbo.VCustomInvoiceAuth
                        WHERE DT >= @FromDate AND DT <= @ToDate";
            var parameters = new DynamicParameters();
            parameters.Add("@FromDate", fromDate.Date);
            parameters.Add("@ToDate", toDate.Date.AddDays(1).AddSeconds(-1));

            if (!string.IsNullOrEmpty(customerCode) && customerCode != "0")
            {
                sql += " AND CustCode = @CustomerCode";
                parameters.Add("@CustomerCode", customerCode);
            }
            if (!string.IsNullOrEmpty(country))
            {
                sql += " AND Country = @Country";
                parameters.Add("@Country", country);
            }
            sql += " ORDER BY DT";
            return await db.QueryAsync<CustomInvoiceAuthItemModel>(sql, parameters);
        }

        // Legacy cmdUpdate_Click: StartTrans ... loop checked items ... con.CommitTrans
        public async Task<int> AuthorizeInvoicesAsync(CustomInvoiceAuthSaveModel request)
        {
            using var db = new SqlConnection(_connectionString);
            await db.OpenAsync();
            using var transaction = db.BeginTransaction();
            try
            {
                int authorizedCount = 0;
                foreach (var item in request.Items)
                {
                    string vchrNo = await AddVoucherAsync(db, transaction, request, item);
                    if (string.IsNullOrEmpty(vchrNo))
                        continue; // Legacy AddVoucher exits without posting when the Pak Rs. amount is zero

                    await db.ExecuteAsync(
                        @"INSERT INTO CustomInvoiceAuthorization(CustomInvoice,VchrNo,AmtInPakRs,ExchRate,UserName,MachineName)
                          VALUES(@CustomInvoice,@VchrNo,@AmtInPakRs,@ExchRate,@UserName,@MachineName)",
                        new
                        {
                            item.CustomInvoice,
                            VchrNo = vchrNo,
                            item.AmtInPakRs,
                            item.ExchRate,
                            request.UserName,
                            request.MachineName
                        }, transaction);

                    authorizedCount++;
                }
                transaction.Commit();
                return authorizedCount;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                _logger.LogError(ex, "Custom invoice authorization failed for user {UserName}. Error: {ErrorMessage}", request.UserName, ex.Message);
                throw;
            }
        }

        // Direct port of the legacy AddVoucher() function (frmCustomInvoiceAuth)
        private async Task<string> AddVoucherAsync(IDbConnection db, IDbTransaction transaction, CustomInvoiceAuthSaveModel request, CustomInvoiceAuthItemModel item)
        {
            double dAmtF = item.TotalCustomAmt;     // Val(.ListSubItems(5).Tag)
            double dExchRate = item.ExchRate;       // Val(.SubItems(6))
            double dAmt = item.AmtInPakRs;          // Val(.SubItems(7))

            if (dAmt == 0) return string.Empty;

            // Advance payments already adjusted against this invoice
            var advances = (await db.QueryAsync<CustomInvoiceAdvanceAdjustmentModel>(
                @"SELECT CustomInvoice, FCustAdvancePaymentsAdjusted.Amount, ExchRate
                  FROM FCustAdvancePayments
                  INNER JOIN FCustAdvancePaymentsAdjusted ON FCustAdvancePayments.EntryID = FCustAdvancePaymentsAdjusted.RefID
                  WHERE CustomInvoice = @CustomInvoice",
                new { item.CustomInvoice }, transaction)).ToList();

            double dTotalAdvAmt = 0;
            double dTotalAdvAmtPK = 0;
            foreach (var advance in advances)
            {
                dTotalAdvAmt += advance.Amount;
                dTotalAdvAmtPK += advance.Amount * advance.ExchRate;
                dTotalAdvAmtPK = Math.Round(dTotalAdvAmtPK);
            }

            DateTime dated = request.VoucherDate.Date;  // DTVchr
            string vchrNo = await _vouchersDataAccess.GetNextVchrNo(dated, "JV");

            // Legacy description: InvNo,Amt @ Rate,Customer Country
            string theDesc = BuildVoucherDescription(item);

            long sNo = await _vouchersDataAccess.GetNextSNo(dated, "Vouchers", db, transaction);

            var voucher = new VoucherViewModel
            {
                VchrNo = vchrNo,
                DT = dated,
                Notes = theDesc,
                UserName = request.UserName,
                MachineName = request.MachineName,
                PostedThroughJVForm = false
            };

            // Charge to Foreign Sales Account ... full amount (Credit entry)
            voucher.LineItems.Add(new VoucherLineItemViewModel
            {
                SNo = sNo++,
                VDate = dated,
                VchrNo = vchrNo,
                AccNo = request.SalesAccNo,
                AccTitle = string.Empty,
                Description = theDesc,
                Debit = 0,
                Credit = (decimal)dAmt,
                Balance = 0,
                CSNo = 0
            });

            // Charge to Customer Account ... amount depends on advance postings against the invoice (Debit entry)
            double dCustomerAmt;
            if (dTotalAdvAmt == 0)
            {
                dCustomerAmt = dAmt;
            }
            else
            {
                dCustomerAmt = dExchRate * (dAmtF - dTotalAdvAmt);
                dCustomerAmt += dTotalAdvAmtPK;
                dCustomerAmt = Math.Round(dCustomerAmt);
            }
            
            voucher.LineItems.Add(new VoucherLineItemViewModel
            {
                SNo = sNo++,
                VDate = dated,
                VchrNo = vchrNo,
                AccNo = item.AccNo,
                AccTitle = string.Empty,
                Description = theDesc,
                Debit = (decimal)dCustomerAmt,
                Credit = 0,
                Balance = 0,
                CSNo = 0
            });

            // Foreign exchange gain/loss when advances were involved
            if (dTotalAdvAmt != 0)
            {
                double dGLCreditAmt = 0;
                double dGLDebitAmt = 0;
                if (dCustomerAmt > dAmt)
                {
                    dGLCreditAmt = dCustomerAmt - dAmt;
                }
                else if (dAmt > dCustomerAmt)
                {
                    dGLDebitAmt = dAmt - dCustomerAmt;
                }

                if (dGLCreditAmt > 0 || dGLDebitAmt > 0)
                {
                    dGLCreditAmt = Math.Round(dGLCreditAmt);
                    dGLDebitAmt = Math.Round(dGLDebitAmt);

                    voucher.LineItems.Add(new VoucherLineItemViewModel
                    {
                        SNo = sNo++,
                        VDate = dated,
                        VchrNo = vchrNo,
                        AccNo = request.ExchDiffAccNo,
                        AccTitle = string.Empty,
                        Description = theDesc,
                        Debit = (decimal)dGLDebitAmt,
                        Credit = (decimal)dGLCreditAmt,
                        Balance = 0,
                        CSNo = 0
                    });
                }
            }

            // Save the voucher using the centralized method
            await _vouchersDataAccess.ExecuteVoucherSave(voucher, db, transaction);

            return vchrNo;
        }

        // Legacy: .SubItems(3) & "," & .SubItems(5) & " @ " & .SubItems(6) & "," & .SubItems(1) & " " & .SubItems(2)
        private static string BuildVoucherDescription(CustomInvoiceAuthItemModel item)
        {
            string amtF = item.TotalCustomAmt.ToString("0.##", CultureInfo.InvariantCulture);
            string rate = item.ExchRate.ToString("0.##", CultureInfo.InvariantCulture);
            return $"{item.CustomInvoice},{item.Curr} {amtF} @ {rate},{item.CustCode} {item.Country}";
        }
    }
}
