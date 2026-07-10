using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Export;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Export
{
    public class CustomPaymentDataAccess : ICustomPaymentDataAccess
    {
        private readonly string _connectionString;

        public CustomPaymentDataAccess(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection") 
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
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
                    fc.Curr AS Currency
                FROM FCustPayments f
                JOIN VCustomInvoiceAuthorized v ON f.CustomInvoice = v.CustomInvoice
                JOIN ForeignCustomers fc ON v.CustCode = fc.CustCode
                WHERE f.EntryID = @EntryID";

            var model = await db.QueryFirstOrDefaultAsync<PostPrcModel>(sql, new { EntryID = entryId });
            return model ?? new PostPrcModel();
        }

        public async Task<string> PostPrcAsync(PostPrcModel model)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            
            // Generate pseudo voucher number for now until full voucher logic is ported
            string vchrNo = "PRC-" + DateTime.Now.ToString("yyyyMMddHHmmss");

            string sql = @"
                INSERT INTO PRC (
                    CustomInvoice, RunningSerialNo, SerialNoDT, BillNo, BillNoDT, 
                    RealizationDT, AmtRealized, ExchRate, UserName, MachineName, VchrNo
                ) VALUES (
                    @CustomInvoice, @SerialNo, @SerialDate, @BillNo, @BillDate, 
                    @RealizationDate, @AmountRealized, @ExchRate, 'Auto', 'Auto', @VchrNo
                );

                UPDATE FCustPayments SET PRCVchrNo = @VchrNo WHERE EntryID = @EntryID;
            ";

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
                VchrNo = vchrNo,
                model.EntryID
            });

            return vchrNo;
        }
    }
}
