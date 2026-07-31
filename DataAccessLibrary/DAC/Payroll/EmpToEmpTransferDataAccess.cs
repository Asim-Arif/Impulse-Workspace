using Dapper;
using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Payroll
{
    public class EmpToEmpTransferDataAccess : IEmpToEmpTransferDataAccess
    {
        private readonly string _connectionString;

        public EmpToEmpTransferDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<EmpTransferBalanceSummaryModel> GetEmployeeBalancesAsync(string empId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            // Outstanding LT Balance (Type = 1)
            const string ltSql = @"
                SELECT ISNULL(SUM(a.Amount), 0) -
                       ISNULL((SELECT SUM(AmtClrd) FROM AmtCleared WHERE EmpID = @EmpID), 0)
                FROM Advances a
                WHERE a.EmpID = @EmpID AND a.Type = 1";

            decimal ltBalance = await db.ExecuteScalarAsync<decimal>(ltSql, new { EmpID = empId });

            // Outstanding ST Balance (Type = 0)
            const string stSql = @"
                SELECT ISNULL(SUM(a.Amount - ISNULL(a.AmountCleared, 0)), 0)
                FROM Advances a
                WHERE a.EmpID = @EmpID AND a.Type = 0 AND (a.Cleared IS NULL OR a.Cleared = 0)";

            decimal stBalance = await db.ExecuteScalarAsync<decimal>(stSql, new { EmpID = empId });

            return new EmpTransferBalanceSummaryModel
            {
                LTBalance = ltBalance,
                STBalance = stBalance
            };
        }

        public async Task ExecuteEmpToEmpTransferAsync(EmpToEmpTransferDto dto, string userName, string machineName)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            if (db.State != ConnectionState.Open) db.Open();
            using IDbTransaction trans = db.BeginTransaction();

            try
            {
                string vchrRefNote = $"Transferred from {dto.FromEmpID}";

                // ── 1. Long Term Loan Transfer ────────────────────────────────
                if (dto.LTAmt > 0)
                {
                    // Add new LT Advance for Destination Employee (Type = 1)
                    const string insertLTAdvSql = @"
                        INSERT INTO Advances
                            (EmpID, DT, Type, Description, Amount, DAmount, AmountCleared, AccVoucherNo, DeductionStartDT)
                        VALUES
                            (@EmpID, @DT, 1, @Description, @Amount, @DAmount, 0, @AccVoucherNo, @DeductionStartDT)";

                    await db.ExecuteAsync(insertLTAdvSql, new
                    {
                        EmpID            = dto.ToEmpID,
                        DT               = dto.TransferDate,
                        Description      = dto.Description,
                        Amount           = dto.LTAmt,
                        DAmount          = dto.LTDeduct,
                        AccVoucherNo     = vchrRefNote,
                        DeductionStartDT = dto.DeductionStartDT
                    }, transaction: trans);

                    // Add AmtCleared record for Source Employee
                    const string insertAmtClearedSql = @"
                        INSERT INTO AmtCleared
                            (EmpID, DT, AmtClrd, VchrNo, Description)
                        VALUES
                            (@EmpID, @DT, @AmtClrd, @VchrNo, @Description)";

                    await db.ExecuteAsync(insertAmtClearedSql, new
                    {
                        EmpID       = dto.FromEmpID,
                        DT          = dto.TransferDate,
                        AmtClrd     = dto.LTAmt,
                        VchrNo      = vchrRefNote,
                        Description = dto.Description
                    }, transaction: trans);
                }

                // ── 2. Short Term Advance Transfer ────────────────────────────
                if (dto.STAmt > 0)
                {
                    // Add new ST Advance for Destination Employee (Type = 0)
                    const string insertSTAdvSql = @"
                        INSERT INTO Advances
                            (EmpID, DT, Type, Description, Amount, DAmount, AmountCleared, AccVoucherNo, DeductionStartDT, ActualAmountTaken)
                        VALUES
                            (@EmpID, @DT, 0, @Description, @Amount, 0, 0, @AccVoucherNo, @DeductionStartDT, @Amount);
                        SELECT CAST(SCOPE_IDENTITY() as bigint);";

                    long lAdvToEntryID = await db.ExecuteScalarAsync<long>(insertSTAdvSql, new
                    {
                        EmpID            = dto.ToEmpID,
                        DT               = dto.TransferDate,
                        Description      = dto.Description,
                        Amount           = dto.STAmt,
                        AccVoucherNo     = vchrRefNote,
                        DeductionStartDT = dto.DeductionStartDT
                    }, transaction: trans);

                    // Add ShortTermAdvanceClearance Header (Type = 1 for Emp to Emp Transfer)
                    const string insertSTACSql = @"
                        INSERT INTO ShortTermAdvanceClearance
                            (EmpID, DT, Remarks, UserName, MachineName, Type, EmpIDTo)
                        VALUES
                            (@EmpID, @DT, @Remarks, @UserName, @MachineName, 1, @EmpIDTo);
                        SELECT CAST(SCOPE_IDENTITY() as bigint);";

                    long lSTACEntryID = await db.ExecuteScalarAsync<long>(insertSTACSql, new
                    {
                        EmpID       = dto.FromEmpID,
                        DT          = dto.TransferDate,
                        Remarks     = dto.Description,
                        UserName    = userName,
                        MachineName = machineName,
                        EmpIDTo     = dto.ToEmpID
                    }, transaction: trans);

                    // Fetch open short term advance rows for Source Employee
                    const string selectOpenSTSql = @"
                        SELECT EntryID, Amount
                        FROM Advances
                        WHERE EmpID = @EmpID AND Type = 0 AND (Cleared IS NULL OR Cleared = 0)
                        ORDER BY Amount DESC";

                    var openAdvances = (await db.QueryAsync<AdvanceItemTemp>(selectOpenSTSql, new { EmpID = dto.FromEmpID }, transaction: trans)).ToList();

                    decimal lAmtLeft = dto.STAmt;

                    foreach (var adv in openAdvances)
                    {
                        decimal lTempAmt = 0;
                        if (adv.Amount >= lAmtLeft)
                        {
                            lTempAmt = lAmtLeft;
                            lAmtLeft = 0;
                        }
                        else
                        {
                            lTempAmt = adv.Amount;
                            lAmtLeft -= lTempAmt;
                        }

                        if (adv.Amount == lTempAmt)
                        {
                            // Fully cleared via transfer
                            await db.ExecuteAsync("UPDATE Advances SET Cleared = 4 WHERE EntryID = @EntryID", new { EntryID = adv.EntryID }, transaction: trans);
                        }
                        else
                        {
                            // Partially cleared via transfer
                            decimal remaining = adv.Amount - lTempAmt;
                            string softDesc = $"Amount Adjusted for Emp. to Emp. Transfer, Original Amt={adv.Amount}, Amt Transferred={lTempAmt} Remain={remaining}";
                            await db.ExecuteAsync(@"
                                UPDATE Advances
                                SET Amount = @Remaining, SoftwareDescription = @SoftDesc
                                WHERE EntryID = @EntryID", new { Remaining = remaining, SoftDesc = softDesc, EntryID = adv.EntryID }, transaction: trans);
                        }

                        // Insert Detail row
                        const string insertSTACDetailSql = @"
                            INSERT INTO ShortTermAdvanceClearanceDetail
                                (SAC_RefID, Advances_RefID, Advances_To_RefID, Amount)
                            VALUES
                                (@SAC_RefID, @Advances_RefID, @Advances_To_RefID, @Amount)";

                        await db.ExecuteAsync(insertSTACDetailSql, new
                        {
                            SAC_RefID         = lSTACEntryID,
                            Advances_RefID    = adv.EntryID,
                            Advances_To_RefID = lAdvToEntryID,
                            Amount            = lTempAmt
                        }, transaction: trans);

                        if (lAmtLeft <= 0) break;
                    }
                }

                trans.Commit();
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }

        private class AdvanceItemTemp
        {
            public long EntryID { get; set; }
            public decimal Amount { get; set; }
        }
    }
}
