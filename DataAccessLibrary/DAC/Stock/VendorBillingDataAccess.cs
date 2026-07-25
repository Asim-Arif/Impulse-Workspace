using Dapper;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Stock
{
    public class VendorBillingDataAccess : IVendorBillingDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<VendorBillingDataAccess> _logger;

        public VendorBillingDataAccess(IConfiguration configuration, ILogger<VendorBillingDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                             ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        public async Task<IEnumerable<VendorLookupModel>> GetVendorsAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            return await db.QueryAsync<VendorLookupModel>(
                "SELECT AccNo, AccTitle FROM VVenders WHERE Active = 1 ORDER BY AccTitle");
        }

        public async Task<IEnumerable<AccountLookupModel>> GetActiveAccountsAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            return await db.QueryAsync<AccountLookupModel>(
                "SELECT AccNo, AccTitle FROM VActiveAccounts ORDER BY AccTitle");
        }

        public async Task<IEnumerable<VendorBillingRowViewModel>> GetUnpostedReceivingsAsync(
            string vendorAccNo, DateTime dateFrom, DateTime dateTo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            var sql = @"
                SELECT EntryID, PORefNo, RcvDate, MaterialID, RMName,
                       QtyRcvd, ExtraQtyRcvd, QtyRejected, Rate, IssuanceType
                FROM VVendRcvdDetailPO
                WHERE RcvDate BETWEEN @DateFrom AND @DateTo
                  AND VendID = @VendorAccNo
                  AND SampleOrder = 0
                  AND EntryID NOT IN (SELECT VRDPO_RefID FROM VenderPostedBillsDetail)
                  AND EntryID NOT IN (SELECT VRDPO_RefID FROM VendRcvdDetailPO_Hidden_Billing)
                ORDER BY RcvDate, EntryID";
            return await db.QueryAsync<VendorBillingRowViewModel>(sql,
                new { DateFrom = dateFrom, DateTo = dateTo, VendorAccNo = vendorAccNo });
        }

        public async Task<IEnumerable<MaterialDeductionRowViewModel>> GetMaterialIssuancesAsync(string vendorAccNo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            var sql = @"
                SELECT EntryID, IssNo, DT, RMID1, RMName, Price, Qty
                FROM VRawMaterialIssuance
                WHERE AccNo = @AccNo AND VchrNo IS NULL
                ORDER BY DT";
            return await db.QueryAsync<MaterialDeductionRowViewModel>(sql, new { AccNo = vendorAccNo });
        }

        public async Task<decimal> GetShortTermDeductionAsync(string vendorAccNo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            var result = await db.QueryFirstOrDefaultAsync<dynamic>(
                "SELECT Amount, AmountCleared FROM VVenderAdvancesShortDeduction WHERE AccNo = @AccNo",
                new { AccNo = vendorAccNo });
            if (result == null) return 0;
            return (decimal)(result.Amount ?? 0) - (decimal)(result.AmountCleared ?? 0);
        }

        public async Task<decimal> GetLongTermDeductionAsync(string vendorAccNo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            var result = await db.QueryFirstOrDefaultAsync<dynamic>(
                "SELECT ActualDeductionAmount FROM VMakerAdvancesDeduction WHERE AccNo = @AccNo",
                new { AccNo = vendorAccNo });
            return result == null ? 0 : (decimal)(result.ActualDeductionAmount ?? 0);
        }

        public async Task<int> GetNextBillNoAsync(string vendorAccNo)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            var max = await db.ExecuteScalarAsync<int?>(
                "SELECT MAX(BillNo) FROM VenderPostedBills WHERE AccNo = @AccNo",
                new { AccNo = vendorAccNo });
            return (max ?? 0) + 1;
        }

        public async Task UpdateReceivingRateAsync(int entryId, float newRate)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            await db.ExecuteAsync(
                "UPDATE VendRcvdDetailPO SET Rate = @Rate WHERE EntryID = @EntryID",
                new { Rate = newRate, EntryID = entryId });
        }

        public async Task HideFromBillingAsync(int entryId, string userName, string machineName)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            await db.ExecuteAsync(
                @"INSERT INTO VendRcvdDetailPO_Hidden_Billing (VRDPO_RefID, UserName, MachineName)
                  VALUES (@EntryID, @UserName, @MachineName)",
                new { EntryID = entryId, UserName = userName, MachineName = machineName });
        }

        public async Task SavePostedBillAsync(VendorBillingPostModel model)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            db.Open();
            using var transaction = db.BeginTransaction();
            try
            {
                // 1. Insert header into VenderPostedBills
                var insertBillSql = @"
                    INSERT INTO VenderPostedBills
                        (AccNo, DT, UserName, MachineName, BillNo, ManualDedAmt, STDeduction, LTDeduction,
                         MaterialDedAmt, VchrNo, DTFrom, DTTo, SampleBilling)
                    VALUES
                        (@AccNo, @DT, @UserName, @MachineName, @BillNo, @ManualDedAmt, @STDeduction, @LTDeduction,
                         @MaterialDedAmt, @VchrNo, @DTFrom, @DTTo, 0);
                    SELECT CAST(SCOPE_IDENTITY() AS INT);";

                var billEntryId = await db.ExecuteScalarAsync<int>(insertBillSql, new
                {
                    AccNo = model.VendorAccNo,
                    DT = DateTime.Today,
                    model.UserName,
                    model.MachineName,
                    model.BillNo,
                    ManualDedAmt = (int)model.ManualDedAmt,
                    STDeduction = (int)model.STDeduction,
                    LTDeduction = (int)model.LTDeduction,
                    MaterialDedAmt = (int)model.MaterialDedAmt,
                    model.VchrNo,
                    DTFrom = model.DateFrom,
                    DTTo = model.DateTo
                }, transaction);

                // 2. Insert checked order receiving entries
                foreach (var id in model.CheckedEntryIDs)
                {
                    await db.ExecuteAsync(
                        "INSERT INTO VenderPostedBillsDetail (RefID, VRDPO_RefID) VALUES (@RefID, @VRDPO_RefID)",
                        new { RefID = billEntryId, VRDPO_RefID = id }, transaction);
                }

                // 3. Insert material issuance detail links
                foreach (var id in model.MaterialIssuanceEntryIDs)
                {
                    await db.ExecuteAsync(
                        "INSERT INTO VenderPostedBillsIssuanceDetails (RefID, VRMID_RefID) VALUES (@RefID, @VRMID_RefID)",
                        new { RefID = billEntryId, VRMID_RefID = id }, transaction);
                }

                // 4. Insert manual deductions
                foreach (var ded in model.ManualDeductions)
                {
                    await db.ExecuteAsync(
                        "INSERT INTO VenderPostedBills_ManualDeductions (Description, Amount, VPB_RefID) VALUES (@Desc, @Amount, @RefID)",
                        new { Desc = ded.Description, Amount = (int)ded.Amount, RefID = billEntryId }, transaction);
                }

                // 5. If Long Term Deduction > 0, record maker amount cleared
                if (model.LTDeduction > 0)
                {
                    await db.ExecuteAsync(
                        "INSERT INTO MakerAmtCleared (AccNo, DT, AmtClrd, VchrNo) VALUES (@AccNo, @DT, @Amt, @VchrNo)",
                        new { AccNo = model.VendorAccNo, DT = model.PostingDate, Amt = (int)model.LTDeduction, VchrNo = model.VchrNo },
                        transaction);
                }

                // 6. If Short Term Deduction > 0, record it
                if (model.STDeduction > 0)
                {
                    await db.ExecuteAsync(
                        "INSERT INTO VenderPostedBills_STDeductions (VPB_RefID, Amount) VALUES (@RefID, @Amount)",
                        new { RefID = billEntryId, Amount = (int)model.STDeduction }, transaction);
                }

                transaction.Commit();
            }
            catch
            {
                transaction.Rollback();
                throw;
            }
        }
    }
}
