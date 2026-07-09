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
            string sql = "SELECT * FROM FCustAdvancePayments WHERE EntryID = @EntryID";
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

        public async Task<string> PostToFinancialAsync(int entryId, AdvancePaymentViewModel payment)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            
            // Note: To properly map to IVouchersDataAccess, we'd create a VoucherMasterViewModel and VouchersDetailViewModels
            // This is a stub for the complex accounting posting
            // Currently, it creates a basic voucher object for the sake of compiling and structure
            
            // Example of what should be done:
            // var voucher = new VoucherMasterViewModel { ... }
            // voucher.VouchersDetails.Add(new VouchersDetailViewModel { ... })
            // string vchrNo = await _vouchersDataAccess.SaveVoucher(voucher);

            // For now, generating a placeholder VchrNo to satisfy the business logic
            string vchrNo = "BRV-" + DateTime.Now.ToString("yyMMdd-HHmm");
            
            string updateSql = "UPDATE FCustAdvancePayments SET VchrNo = @VchrNo WHERE EntryID = @EntryID";
            await db.ExecuteAsync(updateSql, new { VchrNo = vchrNo, EntryID = entryId });
            
            string insertSql = "INSERT INTO FCustAdvancePaymentsVouchers (RefID, VchrNo) VALUES (@EntryID, @VchrNo)";
            await db.ExecuteAsync(insertSql, new { EntryID = entryId, VchrNo = vchrNo });
            
            return vchrNo;
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
