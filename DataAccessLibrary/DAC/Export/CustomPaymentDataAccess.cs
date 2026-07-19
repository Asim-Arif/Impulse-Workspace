using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels.Export;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Export
{
    public class CustomPaymentDataAccess : ICustomPaymentDataAccess
    {
        private readonly string _connectionString;
        private readonly IVouchersDataAccess _vouchersDataAccess;

        public CustomPaymentDataAccess(IConfiguration configuration, IVouchersDataAccess vouchersDataAccess)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection") 
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _vouchersDataAccess = vouchersDataAccess;
        }

        public async Task<List<CustomPaymentStatusModel>> GetCustomPaymentStatusesAsync(string? custCode, int statusIndex)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            var parameters = new DynamicParameters();
            
            if (custCode == "0" || string.IsNullOrWhiteSpace(custCode))
            {
                parameters.Add("@CustCode", null);
            }
            else
            {
                parameters.Add("@CustCode", custCode);
            }

            parameters.Add("@Status", statusIndex);
            parameters.Add("@CustomPayments", 1); // As per VB6 default

            var result = await db.QueryAsync<CustomPaymentStatusModel>(
                "SP_PaymentStatus",
                parameters,
                commandType: CommandType.StoredProcedure);

            return result.ToList();
        }

        public async Task<List<AllCustomersBalanceModel>> GetAllCustomersBalanceAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = @"
                SELECT Curr AS Currency, SUM(TotalCustomAmtToRcv) AS Balance 
                FROM VrptCustomInvoice 
                WHERE FreeSamples=0 
                GROUP BY Curr";

            var balances = await db.QueryAsync<AllCustomersBalanceModel>(sql);
            
            foreach(var balance in balances)
            {
                balance.ExchangeRate = await GetCurrencyExchangeRateAsync(balance.Currency);
            }

            return balances.ToList();
        }

        public async Task<decimal> GetCurrencyExchangeRateAsync(string currency)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = "SELECT ISNULL(MAX(ExchRate), 1) FROM VCurrencyExchangeRates WHERE Currency=@Currency";
            return await db.ExecuteScalarAsync<decimal>(sql, new { Currency = currency });
        }

        public async Task<List<GenericDropDownModel>> GetBanksAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = "SELECT CAST(BankID AS VARCHAR(50)) AS DropDownValue_ID, Bank AS DropDownValue_Description FROM BankList ORDER BY Bank";
            return (await db.QueryAsync<GenericDropDownModel>(sql)).ToList();
        }

        public async Task<List<GenericDropDownModel>> GetPrcBanksAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = "SELECT AccNo AS DropDownValue_ID, AccTitle AS DropDownValue_Description FROM VBankAccounts ORDER BY AccTitle";
            return (await db.QueryAsync<GenericDropDownModel>(sql)).ToList();
        }

        public async Task<List<string>> GetCustomInvoicesForPaymentAsync(string custCode)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = @"
                SELECT CustomInvoice 
                FROM CustomInvoice 
                WHERE CustCode=@CustCode 
                ORDER BY CustomInvoice DESC";

            var result = await db.QueryAsync<string>(sql, new { CustCode = custCode });
            return result.ToList();
        }

        public async Task<decimal> GetCustomInvoiceBalanceAsync(string customInvoice)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = @"SELECT TOP 1 TotalCustomAmtToRcv FROM VrptCustomInvoice WHERE CustomInvoice = @CustomInvoice";
            return await db.QueryFirstOrDefaultAsync<decimal>(sql, new { CustomInvoice = customInvoice });
        }

        public async Task<bool> SaveReceivePaymentAsync(ReceivePaymentModel payment)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = @"
                INSERT INTO FCustPayments(CustomInvoice, DT, AmtRcvd, Description, CustomPayment, FDBCNo, BankID, ExchRate, AmountEE)
                VALUES (@CustomInvoice, @Date, @Amount, @Description, @IsCustomPayment, @FDBCNo, @BankID, @ExchangeRate, @Amount)";

            int rowsAffected = await db.ExecuteAsync(sql, new
            {
                payment.CustomInvoice,
                payment.Date,
                payment.Amount,
                payment.Description,
                IsCustomPayment = payment.IsCustomPayment ? 1 : 0,
                payment.FDBCNo,
                payment.BankID,
                payment.ExchangeRate
            });

            return rowsAffected > 0;
        }

        /// <summary>
        /// Fetches PRC posting details, including:
        /// - CustomerAccNo from ForeignCustomers.AccNo (the debtors/receivable ledger account)
        /// - AuthorizedExchRate from VCustomInvoiceAuthorized.ExchRate (for Exch.Diff calculation)
        /// </summary>
        public async Task<PostPrcModel> GetPrcDetailsAsync(int entryId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string sql = @"
                SELECT 
                    f.EntryID,
                    f.CustomInvoice,
                    v.CustCode,
                    v.Country,
                    f.ExchRate,
                    f.AmtRcvd AS AmountRealized,
                    fc.Curr AS Currency,
                    fc.AccNo AS CustomerAccNo,
                    v.ExchRate AS AuthorizedExchRate
                FROM FCustPayments f
                JOIN VCustomInvoiceAuthorized v ON f.CustomInvoice = v.CustomInvoice
                JOIN ForeignCustomers fc ON v.CustCode = fc.CustCode AND v.Country = fc.Country
                WHERE f.EntryID = @EntryID";

            var model = await db.QueryFirstOrDefaultAsync<PostPrcModel>(sql, new { EntryID = entryId });
            return model ?? new PostPrcModel();
        }

        /// <summary>
        /// Posts the PRC to the financial ledger, building a real voucher matching the
        /// VB6 frmPRC.AddVoucher function (VoucherType = "JV").
        ///
        /// Entries posted:
        ///   DR  Bank Account          = Net Amount (AmtInPakRs - deductions)
        ///   DR  Each deduction acct   = deduction amount (foreign deductions × ExchRate)
        ///   DR  Exch. Diff. account   = Abs(ExchDiff) if ExchDiff is Dr (negative)
        ///   CR  Customer Account      = dAmtForED (AuthorizedExchRate × AmountRealized)
        ///   CR  Exch. Diff. account   = ExchDiff if ExchDiff is Cr (positive)
        /// </summary>
        public async Task<string> PostPrcAsync(PostPrcModel model)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            db.Open();
            using IDbTransaction transaction = db.BeginTransaction();
            try
            {
                // 1. Get next voucher number (VoucherType = "JV" as per VB6 AddVoucher)
                string vchrNo = await _vouchersDataAccess.GetNextVchrNo(model.PostingDate, "JV");

                // 2. Build description string (matching VB6 format)
                string desc = $"{model.CustomInvoice} {model.AmountRealized} {model.Currency} @ {model.ExchRate} {model.BillNo} {model.CustCode}({model.Country})";

                // 3. Calculate amounts
                // Total deductions in PKR (foreign currency deductions × ExchRate)
                decimal dTotalDeductions = 0;
                foreach (var d in model.Deductions)
                {
                    decimal dAmt = d.Currency != "PKR"
                        ? Math.Round(d.Amount * model.ExchRate, 2)
                        : d.Amount;
                    dTotalDeductions += dAmt;
                }

                decimal dNetAmt = Math.Round(model.AmountInRs, 2) - Math.Round(dTotalDeductions, 2);

                // dAmtForED: The amount originally credited to Customer = AuthorizedExchRate × AmountRealized
                decimal dAmtForED = Math.Round(model.AuthorizedExchRate * model.AmountRealized, 4);

                // ExchDiff: positive = Credit side, negative = Debit side
                decimal dExchDiff = Math.Round(model.AmountInRs - dAmtForED, 4);

                // 4. Get next SNo for line items (within same transaction to avoid gaps)
                long sNo = await _vouchersDataAccess.GetNextSNo(model.PostingDate, "Vouchers", db, transaction);

                // 5. Build VoucherViewModel
                var voucher = new VoucherViewModel
                {
                    VchrNo = vchrNo,
                    DT = model.PostingDate,
                    Notes = desc,
                    UserName = model.UserName,
                    MachineName = model.MachineName,
                    PostedThroughJVForm = false
                };

                // DR Bank Account (Net Amount)
                voucher.LineItems.Add(new VoucherLineItemViewModel
                {
                    SNo = sNo++,
                    VDate = model.PostingDate,
                    VchrNo = vchrNo,
                    AccNo = model.BankAccNo,
                    AccTitle = string.Empty,
                    Description = desc,
                    Debit = dNetAmt,
                    Credit = 0,
                    Balance = 0,
                    CSNo = 0
                });

                // DR each deduction account (amount > 0)
                foreach (var d in model.Deductions)
                {
                    if (string.IsNullOrEmpty(d.AccountNo)) continue;
                    decimal dAmt = d.Currency != "PKR"
                        ? Math.Round(d.Amount * model.ExchRate, 2)
                        : d.Amount;
                    dAmt = Math.Round(dAmt, 2);
                    if (dAmt > 0)
                    {
                        voucher.LineItems.Add(new VoucherLineItemViewModel
                        {
                            SNo = sNo++,
                            VDate = model.PostingDate,
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

                // DR Exch. Diff. account if ExchDiff is Dr side (negative difference)
                if (!string.IsNullOrEmpty(model.ExchDiffAccNo) && dExchDiff < 0)
                {
                    voucher.LineItems.Add(new VoucherLineItemViewModel
                    {
                        SNo = sNo++,
                        VDate = model.PostingDate,
                        VchrNo = vchrNo,
                        AccNo = model.ExchDiffAccNo,
                        AccTitle = "Exch. Diff.",
                        Description = desc,
                        Debit = Math.Abs(dExchDiff),
                        Credit = 0,
                        Balance = 0,
                        CSNo = 0
                    });
                }

                // CR Customer Account (dAmtForED = AuthorizedExchRate × AmountRealized)
                voucher.LineItems.Add(new VoucherLineItemViewModel
                {
                    SNo = sNo++,
                    VDate = model.PostingDate,
                    VchrNo = vchrNo,
                    AccNo = model.CustomerAccNo,
                    AccTitle = string.Empty,
                    Description = desc,
                    Debit = 0,
                    Credit = dAmtForED,
                    Balance = 0,
                    CSNo = 0
                });

                // CR Exch. Diff. account if ExchDiff is Cr side (positive difference)
                if (!string.IsNullOrEmpty(model.ExchDiffAccNo) && dExchDiff > 0)
                {
                    voucher.LineItems.Add(new VoucherLineItemViewModel
                    {
                        SNo = sNo++,
                        VDate = model.PostingDate,
                        VchrNo = vchrNo,
                        AccNo = model.ExchDiffAccNo,
                        AccTitle = "Exch. Diff.",
                        Description = desc,
                        Debit = 0,
                        Credit = dExchDiff,
                        Balance = 0,
                        CSNo = 0
                    });
                }

                // 6. Save voucher (within same transaction)
                await _vouchersDataAccess.ExecuteVoucherSave(voucher, db, transaction);

                // 7. Insert into PRC table and update FCustPayments (within same transaction)
                string sql = @"
                    INSERT INTO PRC (
                        CustomInvoice, RunningSerialNo, SerialNoDT, BillNo, BillNoDT,
                        RealizationDT, AmtRealized, ExchRate, UserName, MachineName, VchrNo
                    ) VALUES (
                        @CustomInvoice, @SerialNo, @SerialDate, @BillNo, @BillDate,
                        @RealizationDate, @AmountRealized, @ExchRate, @UserName, @MachineName, @VchrNo
                    );
                    UPDATE FCustPayments SET PRCVchrNo = @VchrNo WHERE EntryID = @EntryID;";

                await db.ExecuteAsync(sql, new
                {
                    model.CustomInvoice,
                    model.SerialNo,
                    model.SerialDate,
                    model.BillNo,
                    model.BillDate,
                    model.RealizationDate,
                    model.AmountRealized,
                    model.ExchRate,
                    model.UserName,
                    model.MachineName,
                    VchrNo = vchrNo,
                    model.EntryID
                }, transaction);

                transaction.Commit();
                return vchrNo;
            }
            catch
            {
                transaction.Rollback();
                throw;
            }
        }
    }
}
