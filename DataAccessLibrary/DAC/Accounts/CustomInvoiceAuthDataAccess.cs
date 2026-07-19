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

        public CustomInvoiceAuthDataAccess(IConfiguration configuration, ILogger<CustomInvoiceAuthDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                            ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
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
            long iNextSno = await GetNextSNoAsync(db, transaction, dated);
            string vchrNo = await db.QueryFirstOrDefaultAsync<string>(
                "SELECT [dbo].[GetNextVchrNo](@DT,@VchrType)",
                new { DT = dated, VchrType = "JV" }, transaction) ?? string.Empty;

            // Legacy description: InvNo,Amt @ Rate,Customer Country
            string theDesc = BuildVoucherDescription(item);

            // Charge to Foreign Sales Account ... full amount (Credit entry)
            string accNo = request.SalesAccNo;
            double bal = await GetBalanceAsync(db, transaction, accNo, dated);
            bal -= dAmt;
            await InsertVoucherLineAsync(db, transaction, iNextSno, dated, vchrNo, accNo, theDesc, 0, dAmt, bal);
            await AdjustLedgerAsync(db, transaction, accNo, dated, iNextSno, -dAmt);

            iNextSno++;

            // Charge to Customer Account ... amount depends on advance postings against the invoice (Debit entry)
            accNo = item.AccNo;
            bal = await GetBalanceAsync(db, transaction, accNo, dated);
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
            bal += dCustomerAmt;
            await InsertVoucherLineAsync(db, transaction, iNextSno, dated, vchrNo, accNo, theDesc, dCustomerAmt, 0, bal);
            await AdjustLedgerAsync(db, transaction, accNo, dated, iNextSno, dCustomerAmt);

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
                    bal = await GetBalanceAsync(db, transaction, request.ExchDiffAccNo, dated);
                    bal += dGLDebitAmt - dGLCreditAmt;

                    iNextSno++;
                    await InsertVoucherLineAsync(db, transaction, iNextSno, dated, vchrNo, request.ExchDiffAccNo, theDesc, dGLDebitAmt, dGLCreditAmt, bal);
                    await AdjustLedgerAsync(db, transaction, request.ExchDiffAccNo, dated, iNextSno, dGLDebitAmt - dGLCreditAmt);
                }
            }

            // Legacy: GetServerDate(True) is the server date with time -> GETDATE()
            await db.ExecuteAsync(
                "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT) VALUES(@VchrNo,@UserName,@MachineName,GETDATE())",
                new { VchrNo = vchrNo, request.UserName, request.MachineName }, transaction);

            return vchrNo;
        }

        // Legacy: .SubItems(3) & "," & .SubItems(5) & " @ " & .SubItems(6) & "," & .SubItems(1) & " " & .SubItems(2)
        private static string BuildVoucherDescription(CustomInvoiceAuthItemModel item)
        {
            string amtF = item.TotalCustomAmt.ToString("0.##", CultureInfo.InvariantCulture);
            string rate = item.ExchRate.ToString("0.##", CultureInfo.InvariantCulture);
            return $"{item.CustomInvoice},{item.Curr} {amtF} @ {rate},{item.CustCode} {item.Country}";
        }

        // Direct port of the legacy getNextSno() function (Vouchers table, per voucher date)
        private async Task<long> GetNextSNoAsync(IDbConnection db, IDbTransaction transaction, DateTime onDate)
        {
            long? maxSNo = await db.QueryFirstOrDefaultAsync<long?>(
                "SELECT MAX(SNo) AS MaxSNo FROM Vouchers WHERE VDate=@VDate",
                new { VDate = onDate }, transaction);
            if (maxSNo == null)
                return long.Parse(onDate.ToString("yyMMdd") + "0001");
            return (maxSNo ?? 1) + 1;
        }

        // Legacy getBalance() is mirrored by the dbo.GetBalance scalar function (same as VouchersDataAccess.GetBalance)
        private async Task<double> GetBalanceAsync(IDbConnection db, IDbTransaction transaction, string accNo, DateTime onDate)
        {
            return await db.QueryFirstOrDefaultAsync<double>(
                "SELECT [dbo].[GetBalance](@AccNo, @DT, @CurrentDT)",
                new { AccNo = accNo, DT = onDate, CurrentDT = DateTime.Now }, transaction);
        }

        // Direct port of the legacy adjustLedger() sub (Vouchers.bas)
        private async Task AdjustLedgerAsync(IDbConnection db, IDbTransaction transaction, string accNo, DateTime startDate, long sNo, double amountToAdd)
        {
            await db.ExecuteAsync(
                "UPDATE Accounts SET Balance = Balance + @Amount WHERE AccNo = @AccNo",
                new { Amount = amountToAdd, AccNo = accNo }, transaction);
            await db.ExecuteAsync(
                "UPDATE Vouchers SET Balance = Balance + @Amount WHERE AccNo = @AccNo AND (VDate >= @StartDate AND SNo > @SNo)",
                new { Amount = amountToAdd, AccNo = accNo, StartDate = startDate, SNo = sNo }, transaction);
        }

        // Legacy: INSERT INTO Vouchers VALUES(SNo,Dated,VchrNo,AccNo,Desc,Debit,Credit,Bal,'',0)
        private async Task InsertVoucherLineAsync(IDbConnection db, IDbTransaction transaction, long sNo, DateTime vDate, string vchrNo, string accNo, string description, double debit, double credit, double balance)
        {
            await db.ExecuteAsync(
                @"INSERT INTO Vouchers(SNo,VDate,VchrNo,Accno,Description,Debit,Credit,balance,DpstSlip,CSNo)
                  VALUES(@SNo,@VDate,@VchrNo,@Accno,@Description,@Debit,@Credit,@Balance,'',0)",
                new
                {
                    SNo = sNo,
                    VDate = vDate,
                    VchrNo = vchrNo,
                    Accno = accNo,
                    Description = description,
                    Debit = debit,
                    Credit = credit,
                    Balance = balance
                }, transaction);
        }
    }
}
