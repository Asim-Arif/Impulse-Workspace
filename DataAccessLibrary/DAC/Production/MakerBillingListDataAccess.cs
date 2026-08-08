using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Production
{
    public class MakerBillingListDataAccess : IMakerBillingListDataAccess
    {
        private readonly string ConnectionString;

        public MakerBillingListDataAccess(IConfiguration configuration)
        {
            ConnectionString = configuration.GetConnectionString("DefaultConnection")
                ?? configuration.GetConnectionString("SqlConnectionString")
                ?? string.Empty;
        }

        public async Task<List<LookupItemInt>> GetMakersAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT VendID AS Id,
                       CASE WHEN ISNULL(VendID1, '') = '' THEN VenderName ELSE '[' + VendID1 + '] ' + VenderName END AS Name
                FROM Makers
                WHERE ISNULL(Active, 1) = 1
                ORDER BY VenderName";

            var list = (await db.QueryAsync<LookupItemInt>(sql)).ToList();
            list.Insert(0, new LookupItemInt { Id = 0, Name = "<All Makers>" });
            return list;
        }

        public async Task<List<MakerBillingListItem>> GetPostedBillsListAsync(MakerBillingListFilter filter)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            string sql = @"
                SELECT EntryID, VendID, VenderName, VendID1, BillNo, VchrNo, DT,
                       ISNULL(TotalQty, 0) AS TotalQty,
                       ISNULL(TotalValue, 0) AS TotalValue,
                       ISNULL(NetPayable, 0) AS NetPayable,
                       ISNULL(TotalPaidAmount, 0) AS TotalPaidAmount,
                       ISNULL(STDeduction, 0) AS STDeduction,
                       ISNULL(LTDeduction, 0) AS LTDeduction,
                       ISNULL(Posted, 0) AS Posted,
                       Payment_VchrNo
                FROM VMakerPostedBillsList
                WHERE DT >= @DtFrom AND DT <= @DtTo";

            var p = new DynamicParameters();
            p.Add("@DtFrom", filter.DtFrom.Date);
            p.Add("@DtTo", filter.DtTo.Date.AddDays(1).AddTicks(-1));

            if (filter.VendID > 0)
            {
                sql += " AND VendID = @VendID";
                p.Add("@VendID", filter.VendID);
            }

            sql += " ORDER BY EntryID DESC";

            var result = await db.QueryAsync<MakerBillingListItem>(sql, p);
            return result.ToList();
        }

        public async Task<List<LookupItemString>> GetPaymentCreditAccountsAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT AccNo AS Id, '[' + AccNo + '] ' + AccTitle AS Name
                FROM Accounts
                WHERE ISNULL(Active, 1) = 1 AND ISNULL(Parent, 0) = 0
                ORDER BY AccTitle";

            var list = (await db.QueryAsync<LookupItemString>(sql)).ToList();
            return list;
        }

        public async Task<string?> GetMakerAccountNoAsync(int vendId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT AccNo FROM Makers WHERE VendID = @VendID";
            return await db.QueryFirstOrDefaultAsync<string?>(sql, new { VendID = vendId });
        }

        public async Task<bool> RecordPaymentVoucherAsync(long mpbRefId, string voucherNo, decimal paidAmt)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                INSERT INTO MakerPostedBills_Payments (MPB_RefID, VchrNo, PaidAmt)
                VALUES (@MPB_RefID, @VchrNo, @PaidAmt)";

            int rows = await db.ExecuteAsync(sql, new
            {
                MPB_RefID = (int)mpbRefId,
                VchrNo = voucherNo,
                PaidAmt = (float)paidAmt
            });
            return rows > 0;
        }

        public async Task<bool> DeletePostedBillAsync(long entryId, string userName)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            if (db.State != ConnectionState.Open) db.Open();
            using var tx = db.BeginTransaction();

            try
            {
                // 1. Check if bill exists
                const string checkSql = "SELECT VchrNo FROM MakerPostedBills WHERE EntryID = @EntryID";
                string? vchrNo = await db.ExecuteScalarAsync<string?>(checkSql, new { EntryID = entryId }, tx);
                if (string.IsNullOrEmpty(vchrNo))
                {
                    tx.Rollback();
                    return false;
                }

                // 2. Unlink VendIssued
                const string unlinkSql = @"
                    UPDATE VendIssued SET VchrNo = NULL 
                    WHERE EntryID IN (SELECT Iss_RefID FROM MakerPostedBillsDetail WHERE RefID = @EntryID)";
                await db.ExecuteAsync(unlinkSql, new { EntryID = entryId }, tx);

                // 3. Delete details
                await db.ExecuteAsync("DELETE FROM MakerPostedBillsDetail_Receivings WHERE MPB_RefID = @EntryID", new { EntryID = entryId }, tx);
                await db.ExecuteAsync("DELETE FROM MakerPostedBillsDetail WHERE RefID = @EntryID", new { EntryID = entryId }, tx);
                await db.ExecuteAsync("DELETE FROM MakerPostedBills_MaterialDeductions WHERE MPB_RefID = @EntryID", new { EntryID = entryId }, tx);
                await db.ExecuteAsync("DELETE FROM MakerPostedBills_ManualDeductions WHERE MPB_RefID = @EntryID", new { EntryID = entryId }, tx);
                await db.ExecuteAsync("DELETE FROM MakerPostedBills_Payments WHERE MPB_RefID = @EntryID", new { EntryID = entryId }, tx);

                // 4. Delete master
                await db.ExecuteAsync("DELETE FROM MakerPostedBills WHERE EntryID = @EntryID", new { EntryID = entryId }, tx);

                tx.Commit();
                return true;
            }
            catch
            {
                tx.Rollback();
                throw;
            }
        }
    }
}
