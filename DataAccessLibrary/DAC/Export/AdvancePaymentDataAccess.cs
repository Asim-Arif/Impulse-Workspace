using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using Dapper;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using DataAccessLibrary.Models.ViewModels.Export;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels.Accounts;

namespace DataAccessLibrary.DAC.Export
{
    public class AdvancePaymentDataAccess : IAdvancePaymentDataAccess
    {
        private readonly string _connectionString;
        private readonly IVouchersDataAccess _vouchersDataAccess;

        public AdvancePaymentDataAccess(IConfiguration configuration, IVouchersDataAccess vouchersDataAccess)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection") 
                                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _vouchersDataAccess = vouchersDataAccess;
        }

        public async Task<List<AdvancePaymentListViewModel>> GetAdvancePaymentListAsync(DateTime dtFrom, DateTime dtTo, string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            
            var parameters = new DynamicParameters();
            parameters.Add("@DtFrom", dtFrom);
            parameters.Add("@DtTo", dtTo);
            
            string sql = @"
                SELECT 
                    ap.*,
                    c.Name AS CustomerName,
                    b.AccTitle AS BankName
                FROM FCustAdvancePayments ap
                LEFT JOIN ForeignCustomers c ON ap.CustCode = c.CustCode
                LEFT JOIN VBankAccounts b ON ap.BankAccNo = b.AccNo
                WHERE ap.DT >= @DtFrom AND ap.DT <= @DtTo ";

            if (!string.IsNullOrEmpty(custCode) && custCode != "0")
            {
                sql += " AND ap.CustCode = @CustCode ";
                parameters.Add("@CustCode", custCode);
            }
            if (!string.IsNullOrEmpty(country) && country != "<All Countries>")
            {
                sql += " AND ap.Country = @Country ";
                parameters.Add("@Country", country);
            }
            
            sql += " ORDER BY ap.DT DESC";

            var list = await db.QueryAsync<AdvancePaymentListViewModel>(sql, parameters);
            return list.AsList();
        }

        public async Task<AdvancePaymentViewModel?> GetAdvancePaymentAsync(int entryId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            // JOIN ForeignCustomers to also fetch the customer's ledger AccNo
            string sql = @"
                SELECT ap.*, fc.AccNo AS CustomerAccNo
                FROM FCustAdvancePayments ap
                LEFT JOIN ForeignCustomers fc ON ap.CustCode = fc.CustCode AND ap.Country = fc.Country
                WHERE ap.EntryID = @EntryID";
            return await db.QuerySingleOrDefaultAsync<AdvancePaymentViewModel>(sql, new { EntryID = entryId });
        }

        public async Task<bool> SaveAdvancePaymentAsync(AdvancePaymentViewModel payment)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            
            if (payment.EntryID == 0)
            {
                string sql = @"
                    INSERT INTO FCustAdvancePayments (
                        DT, AdviceNo, CustCode, Country, Amount, ImporterAddress, 
                        ExportOf, ExchRate, BankAccNo, UserName, MachineName, OrderNo
                    ) VALUES (
                        @DT, @AdviceNo, @CustCode, @Country, @Amount, @ImporterAddress, 
                        @ExportOf, @ExchRate, @BankAccNo, @UserName, @MachineName, @OrderNo
                    )";
                await db.ExecuteAsync(sql, payment);
            }
            else
            {
                string sql = @"
                    UPDATE FCustAdvancePayments SET 
                        DT = @DT,
                        AdviceNo = @AdviceNo,
                        CustCode = @CustCode,
                        Country = @Country,
                        Amount = @Amount,
                        ImporterAddress = @ImporterAddress,
                        ExportOf = @ExportOf,
                        ExchRate = @ExchRate,
                        BankAccNo = @BankAccNo,
                        OrderNo = @OrderNo
                    WHERE EntryID = @EntryID";
                await db.ExecuteAsync(sql, payment);
            }

            return true;
        }

        public async Task<bool> UpdateExchangeRateAsync(int entryId, decimal exchRate)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = "UPDATE FCustAdvancePayments SET ExchRate = @ExchRate WHERE EntryID = @EntryID";
            await db.ExecuteAsync(sql, new { EntryID = entryId, ExchRate = exchRate });
            return true;
        }

        /// <summary>
        /// Posts the advance payment to the financial ledger by building a real voucher
        /// matching the VB6 frmAdvancePaymentPosting.AddVoucher logic (VoucherType = "ADV").
        /// 
        /// Entries posted:
        ///   DR  Bank Account          = Net Amount (AmtInPakRs - all deductions)
        ///   DR  Each deduction acct   = deduction amount (USD deductions × ExchRate)
        ///   CR  Customer Account      = Full Amount in Rs (AmtInPakRs)
        /// </summary>
        public async Task<string> PostToFinancialAsync(
            int entryId,
            AdvancePaymentViewModel payment,
            List<PrcDeductionModel> deductions,
            DateTime postingDate)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            db.Open();
            using IDbTransaction transaction = db.BeginTransaction();
            try
            {
                // 1. Get next voucher number (VoucherType = "ADV" as per VB6)
                string vchrNo = await _vouchersDataAccess.GetNextVchrNo(postingDate, "ADV");

                // 2. Build description (matching VB6 format)
                decimal amtInPakRs = Math.Round(payment.Amount * payment.ExchRate, 2);
                string desc = $"{payment.AdviceNo} {payment.Amount} @ {payment.ExchRate} {payment.CustCode}({payment.Country})";

                // 3. Calculate total deductions in PKR (USD deductions use ExchRate)
                decimal dTotalDeductions = 0;
                foreach (var d in deductions)
                {
                    decimal amt = d.Currency != "PKR"
                        ? Math.Round(d.Amount * payment.ExchRate, 2)
                        : d.Amount;
                    dTotalDeductions += Math.Round(amt, 2);
                }
                decimal dNetAmt = Math.Round(amtInPakRs - dTotalDeductions, 2);

                // 4. Get next SNo for line items (within same transaction)
                long sNo = await _vouchersDataAccess.GetNextSNo(postingDate, "Vouchers", db, transaction);

                // 5. Build VoucherViewModel
                var voucher = new VoucherViewModel
                {
                    VchrNo = vchrNo,
                    DT = postingDate,
                    Notes = desc,
                    UserName = payment.UserName,
                    MachineName = payment.MachineName,
                    PostedThroughJVForm = false
                };

                // DR Bank Account (Net Amount)
                voucher.LineItems.Add(new VoucherLineItemViewModel
                {
                    SNo = sNo++,
                    VDate = postingDate,
                    VchrNo = vchrNo,
                    AccNo = payment.BankAccNo,
                    AccTitle = string.Empty,
                    Description = desc,
                    Debit = dNetAmt,
                    Credit = 0,
                    Balance = 0,
                    CSNo = 0
                });

                // DR each deduction account with amount > 0
                foreach (var d in deductions)
                {
                    if (string.IsNullOrEmpty(d.AccountNo)) continue;
                    decimal dAmt = d.Currency != "PKR"
                        ? Math.Round(d.Amount * payment.ExchRate, 2)
                        : d.Amount;
                    dAmt = Math.Round(dAmt, 2);
                    if (dAmt > 0)
                    {
                        voucher.LineItems.Add(new VoucherLineItemViewModel
                        {
                            SNo = sNo++,
                            VDate = postingDate,
                            VchrNo = vchrNo,
                            AccNo = d.AccountNo,
                            AccTitle = d.Title,
                            Description = desc,
                            Debit = dAmt,
                            Credit = 0,
                            Balance = 0,
                            CSNo = 0
                        });
                    }
                }

                // CR Customer Account (full Amount in Rs)
                voucher.LineItems.Add(new VoucherLineItemViewModel
                {
                    SNo = sNo++,
                    VDate = postingDate,
                    VchrNo = vchrNo,
                    AccNo = payment.CustomerAccNo,
                    AccTitle = string.Empty,
                    Description = desc,
                    Debit = 0,
                    Credit = amtInPakRs,
                    Balance = 0,
                    CSNo = 0
                });

                // 6. Save voucher in same transaction
                await _vouchersDataAccess.ExecuteVoucherSave(voucher, db, transaction);

                // 7. Update FCustAdvancePayments and insert into vouchers link table
                string updateSql = "UPDATE FCustAdvancePayments SET VchrNo = @VchrNo WHERE EntryID = @EntryID";
                await db.ExecuteAsync(updateSql, new { VchrNo = vchrNo, EntryID = entryId }, transaction);

                string insertSql = "INSERT INTO FCustAdvancePaymentsVouchers (RefID, VchrNo) VALUES (@EntryID, @VchrNo)";
                await db.ExecuteAsync(insertSql, new { EntryID = entryId, VchrNo = vchrNo }, transaction);

                transaction.Commit();
                return vchrNo;
            }
            catch
            {
                transaction.Rollback();
                throw;
            }
        }

        public async Task<bool> DeleteAdvancePaymentAsync(int entryId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = "DELETE FROM FCustAdvancePayments WHERE EntryID = @EntryID";
            await db.ExecuteAsync(sql, new { EntryID = entryId });
            return true;
        }

        public async Task<List<AdvancePaymentBankModel>> GetBanksAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = "SELECT AccNo AS BankAccNo, AccTitle AS BankName FROM VBankAccounts ORDER BY AccTitle";
            var list = await db.QueryAsync<AdvancePaymentBankModel>(sql);
            return list.AsList();
        }

        public async Task<List<string>> GetOrdersAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = "SELECT DISTINCT OrderNo FROM FCustomerOrders WHERE CustCode = @CustCode AND Country = @Country ORDER BY OrderNo";
            var list = await db.QueryAsync<string>(sql, new { CustCode = custCode, Country = country });
            return list.AsList();
        }

        public async Task<(string Currency, string Address)> GetCustomerDetailsAsync(string custCode, string country)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            
            string currSql = "SELECT Curr FROM ForeignCustomers WHERE CustCode = @CustCode AND Country = @Country";
            string currency = await db.QueryFirstOrDefaultAsync<string>(currSql, new { CustCode = custCode, Country = country }) ?? string.Empty;

            string addrSql = @"
                SELECT TOP 1 Address2 
                FROM FCustBanks 
                WHERE CustCode = @CustCode 
                ORDER BY DefaultBank DESC";
            string address = await db.QueryFirstOrDefaultAsync<string>(addrSql, new { CustCode = custCode }) ?? string.Empty;

            return (currency, address);
        }

        public async Task<List<AdvancePaymentUsageViewModel>> GetUsageDetailsAsync(int entryId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = "SELECT * FROM FCustAdvancePaymentsAdjusted WHERE RefID = @RefID ORDER BY EntryID";
            return (await db.QueryAsync<AdvancePaymentUsageViewModel>(sql, new { RefID = entryId })).AsList();
        }
    }
}
