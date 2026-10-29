using Dapper;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

public class BankListDataAccess : IBankListDataAccess
{
    private readonly string _connectionString;

    public BankListDataAccess(IConfiguration config)
    {
        _connectionString = config.GetConnectionString("DefaultConnection");
    }

    // ─────────────────────────────────────────────────────────────
    // Modern Hierarchical Queries
    // ─────────────────────────────────────────────────────────────

    public async Task<List<BankBranchRecord>> GetAllBankBranchesAsync()
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            const string sql = @"
                SELECT 
                    b.BankID,
                    ISNULL(b.Bank, '') AS Bank,
                    ISNULL(b.Branch, '') AS Branch,
                    ISNULL(b.City, '') AS City,
                    ISNULL(b.Phone1, '') AS Phone1,
                    ISNULL(b.Phone2, '') AS Phone2,
                    ISNULL(b.Fax, '') AS Fax,
                    ISNULL(b.Email, '') AS Email,
                    ISNULL(b.Manager, '') AS Manager,
                    ISNULL(b.Address, '') AS Address,
                    ISNULL(b.AccNo, '') AS AccNo,
                    ISNULL(b.Remittance, '') AS Remittance,
                    b.CompanyRefID,
                    COUNT(ba.AccNo) AS AccountsCount
                FROM BankList b
                LEFT JOIN BankAccounts ba ON b.BankID = ba.BankID
                GROUP BY 
                    b.BankID, b.Bank, b.Branch, b.City, b.Phone1, b.Phone2, 
                    b.Fax, b.Email, b.Manager, b.Address, b.AccNo, b.Remittance, b.CompanyRefID
                ORDER BY b.Bank ASC, b.Branch ASC, b.City ASC";

            var rows = await db.QueryAsync<BankBranchRecord>(sql);
            return rows.ToList();
        }
    }

    public async Task<List<BankAccountItem>> GetAllBankAccountsAsync()
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            const string sql = @"
                SELECT 
                    ba.BankID,
                    ISNULL(b.Bank, '') AS Bank,
                    ISNULL(b.Branch, '') AS Branch,
                    ISNULL(b.City, '') AS City,
                    ba.AccNo,
                    ISNULL(a.AccTitle, '') AS AccTitle,
                    ISNULL(ba.BankAccNo, '') AS BankAccNo,
                    ISNULL(ba.Type, 'Current') AS Type,
                    ba.DefaultBank,
                    ba.AllowOD,
                    ba.ODAmt,
                    ba.ODRate,
                    ISNULL(ba.Chqformat, '') AS Chqformat,
                    ISNULL(a.Balance, 0) AS Balance,
                    ISNULL(a.openbal, 0) AS OpenBal,
                    a.opendate AS OpenDate,
                    a.Active,
                    (SELECT COUNT(1) FROM ChqBooks cb WHERE cb.AccNo = ba.AccNo) AS ChqBooksCount
                FROM BankAccounts ba
                INNER JOIN BankList b ON ba.BankID = b.BankID
                INNER JOIN Accounts a ON ba.AccNo = a.AccNo
                ORDER BY b.Bank ASC, b.Branch ASC, ba.DefaultBank DESC, a.AccTitle ASC";

            var rows = await db.QueryAsync<BankAccountItem>(sql);
            return rows.ToList();
        }
    }

    public async Task<List<BankAccountItem>> GetBankAccountsByBankIdAsync(int bankId)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            const string sql = @"
                SELECT 
                    ba.BankID,
                    ISNULL(b.Bank, '') AS Bank,
                    ISNULL(b.Branch, '') AS Branch,
                    ISNULL(b.City, '') AS City,
                    ba.AccNo,
                    ISNULL(a.AccTitle, '') AS AccTitle,
                    ISNULL(ba.BankAccNo, '') AS BankAccNo,
                    ISNULL(ba.Type, 'Current') AS Type,
                    ba.DefaultBank,
                    ba.AllowOD,
                    ba.ODAmt,
                    ba.ODRate,
                    ISNULL(ba.Chqformat, '') AS Chqformat,
                    ISNULL(a.Balance, 0) AS Balance,
                    ISNULL(a.openbal, 0) AS OpenBal,
                    a.opendate AS OpenDate,
                    a.Active,
                    (SELECT COUNT(1) FROM ChqBooks cb WHERE cb.AccNo = ba.AccNo) AS ChqBooksCount
                FROM BankAccounts ba
                INNER JOIN BankList b ON ba.BankID = b.BankID
                INNER JOIN Accounts a ON ba.AccNo = a.AccNo
                WHERE ba.BankID = @BankID
                ORDER BY ba.DefaultBank DESC, a.AccTitle ASC";

            var rows = await db.QueryAsync<BankAccountItem>(sql, new { BankID = bankId });
            return rows.ToList();
        }
    }

    public async Task<BankAccountItem?> GetBankAccountByAccNoAsync(string accNo)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            const string sql = @"
                SELECT 
                    ba.BankID,
                    ISNULL(b.Bank, '') AS Bank,
                    ISNULL(b.Branch, '') AS Branch,
                    ISNULL(b.City, '') AS City,
                    ba.AccNo,
                    ISNULL(a.AccTitle, '') AS AccTitle,
                    ISNULL(ba.BankAccNo, '') AS BankAccNo,
                    ISNULL(ba.Type, 'Current') AS Type,
                    ba.DefaultBank,
                    ba.AllowOD,
                    ba.ODAmt,
                    ba.ODRate,
                    ISNULL(ba.Chqformat, '') AS Chqformat,
                    ISNULL(a.Balance, 0) AS Balance,
                    ISNULL(a.openbal, 0) AS OpenBal,
                    a.opendate AS OpenDate,
                    a.Active,
                    (SELECT COUNT(1) FROM ChqBooks cb WHERE cb.AccNo = ba.AccNo) AS ChqBooksCount
                FROM BankAccounts ba
                INNER JOIN BankList b ON ba.BankID = b.BankID
                INNER JOIN Accounts a ON ba.AccNo = a.AccNo
                WHERE ba.AccNo = @AccNo";

            return await db.QuerySingleOrDefaultAsync<BankAccountItem>(sql, new { AccNo = accNo });
        }
    }

    public async Task<BankBranchRecord?> GetBankBranchByIdAsync(int bankId)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            const string sql = @"
                SELECT 
                    b.BankID,
                    ISNULL(b.Bank, '') AS Bank,
                    ISNULL(b.Branch, '') AS Branch,
                    ISNULL(b.City, '') AS City,
                    ISNULL(b.Phone1, '') AS Phone1,
                    ISNULL(b.Phone2, '') AS Phone2,
                    ISNULL(b.Fax, '') AS Fax,
                    ISNULL(b.Email, '') AS Email,
                    ISNULL(b.Manager, '') AS Manager,
                    ISNULL(b.Address, '') AS Address,
                    ISNULL(b.AccNo, '') AS AccNo,
                    ISNULL(b.Remittance, '') AS Remittance,
                    b.CompanyRefID,
                    COUNT(ba.AccNo) AS AccountsCount
                FROM BankList b
                LEFT JOIN BankAccounts ba ON b.BankID = ba.BankID
                WHERE b.BankID = @BankID
                GROUP BY 
                    b.BankID, b.Bank, b.Branch, b.City, b.Phone1, b.Phone2, 
                    b.Fax, b.Email, b.Manager, b.Address, b.AccNo, b.Remittance, b.CompanyRefID";

            return await db.QuerySingleOrDefaultAsync<BankBranchRecord>(sql, new { BankID = bankId });
        }
    }

    public async Task<List<string>> GetDistinctBankNamesAsync()
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            const string sql = "SELECT DISTINCT Bank FROM BankList WHERE ISNULL(Bank, '') <> '' ORDER BY Bank ASC";
            var rows = await db.QueryAsync<string>(sql);
            return rows.ToList();
        }
    }

    public async Task<List<string>> GetDistinctCitiesAsync()
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            const string sql = "SELECT DISTINCT City FROM BankList WHERE ISNULL(City, '') <> '' ORDER BY City ASC";
            var rows = await db.QueryAsync<string>(sql);
            return rows.ToList();
        }
    }

    // ─────────────────────────────────────────────────────────────
    // Modern Transaction-Backed CRUD Operations
    // ─────────────────────────────────────────────────────────────

    public async Task<BankOperationResult> SaveBankBranchAsync(SaveBankBranchDto dto, string userName, string machineName)
    {
        if (string.IsNullOrWhiteSpace(dto.Bank))
        {
            return new BankOperationResult { Success = false, Message = "Bank Name is required." };
        }
        if (string.IsNullOrWhiteSpace(dto.Branch))
        {
            return new BankOperationResult { Success = false, Message = "Branch Name is required." };
        }

        using (var conn = new SqlConnection(_connectionString))
        {
            await conn.OpenAsync();
            try
            {
                if (dto.BankID > 0)
                {
                    const string updateSql = @"
                        UPDATE BankList 
                        SET Bank = @Bank,
                            Branch = @Branch,
                            City = @City,
                            Phone1 = @Phone1,
                            Phone2 = @Phone2,
                            Fax = @Fax,
                            Email = @Email,
                            Manager = @Manager,
                            Address = @Address,
                            Remittance = @Remittance,
                            CompanyRefID = @CompanyRefID
                        WHERE BankID = @BankID";

                    await conn.ExecuteAsync(updateSql, dto);
                    return new BankOperationResult
                    {
                        Success = true,
                        Message = "Bank branch updated successfully.",
                        BankID = dto.BankID
                    };
                }
                else
                {
                    const string insertSql = @"
                        INSERT INTO BankList 
                        (Bank, Branch, City, Phone1, Phone2, Fax, Email, Manager, Address, Remittance, CompanyRefID)
                        OUTPUT INSERTED.BankID
                        VALUES 
                        (@Bank, @Branch, @City, @Phone1, @Phone2, @Fax, @Email, @Manager, @Address, @Remittance, @CompanyRefID)";

                    int newId = await conn.ExecuteScalarAsync<int>(insertSql, dto);
                    return new BankOperationResult
                    {
                        Success = true,
                        Message = "New bank branch registered successfully.",
                        BankID = newId
                    };
                }
            }
            catch (Exception ex)
            {
                return new BankOperationResult { Success = false, Message = $"Failed to save bank branch: {ex.Message}" };
            }
        }
    }

    public async Task<BankOperationResult> SaveBankAccountAsync(SaveBankAccountDto dto, string userName, string machineName)
    {
        if (dto.BankID <= 0)
        {
            return new BankOperationResult { Success = false, Message = "A valid Bank Branch must be selected." };
        }
        if (string.IsNullOrWhiteSpace(dto.AccTitle))
        {
            return new BankOperationResult { Success = false, Message = "Account Title is required." };
        }
        if (string.IsNullOrWhiteSpace(dto.BankAccNo))
        {
            return new BankOperationResult { Success = false, Message = "Bank Account Number is required." };
        }

        using (var conn = new SqlConnection(_connectionString))
        {
            await conn.OpenAsync();
            using (var tran = conn.BeginTransaction())
            {
                try
                {
                    if (dto.IsNew)
                    {
                        // 1. Locate Master Parent Account from GeneralData or fallback to '15-002'
                        string masterAccNo = await conn.ExecuteScalarAsync<string>(
                            "SELECT TOP 1 DataValue FROM GeneralData WHERE DataName = 'Bank_Master_AccNo'", transaction: tran);

                        if (string.IsNullOrWhiteSpace(masterAccNo))
                        {
                            masterAccNo = "15-002";
                        }

                        // 2. Generate Next Account Number under 15-002 using database procedure
                        string nextAccNo = await conn.ExecuteScalarAsync<string>(
                            "SELECT dbo.GetNextAccno(@AccountName, '15', @ParentAccount, 0)",
                            new { AccountName = dto.AccTitle, ParentAccount = masterAccNo }, transaction: tran);

                        if (string.IsNullOrWhiteSpace(nextAccNo) || nextAccNo.StartsWith("-"))
                        {
                            // Safety fallback if DB function returns malformed string
                            int currentCount = await conn.ExecuteScalarAsync<int>(
                                "SELECT COUNT(1) FROM Accounts WHERE SubAccOf = @Parent",
                                new { Parent = masterAccNo }, transaction: tran);

                            nextAccNo = $"{masterAccNo}-{(currentCount + 1).ToString("D5")}";
                        }

                        dto.AccNo = nextAccNo;

                        // 3. Insert into Accounts (GL Master) with Active=1 and Parent=0
                        const string insertAccSql = @"
                            INSERT INTO Accounts 
                            (AccNo, AccTitle, SubAccOf, Type, openbal, opendate, Balance, Parent, Active, UserName, ComputerName)
                            VALUES 
                            (@AccNo, @AccTitle, @SubAccOf, '15', @OpenBal, @OpenDate, @OpenBal, 0, 1, @UserName, @ComputerName)";

                        await conn.ExecuteAsync(insertAccSql, new
                        {
                            AccNo = dto.AccNo,
                            AccTitle = dto.AccTitle,
                            SubAccOf = masterAccNo,
                            OpenBal = dto.OpenBal,
                            OpenDate = dto.OpenDate ?? DateTime.Today,
                            UserName = userName ?? "System",
                            ComputerName = machineName ?? "Web"
                        }, transaction: tran);

                        // 4. Default bank exclusivity
                        if (dto.DefaultBank)
                        {
                            await conn.ExecuteAsync("UPDATE BankAccounts SET DefaultBank = 0", transaction: tran);
                        }

                        // 5. Insert into BankAccounts
                        const string insertBankAccSql = @"
                            INSERT INTO BankAccounts 
                            (BankID, AccNo, Type, BankAccNo, DefaultBank, AllowOD, ODAmt, ODRate, Chqformat)
                            VALUES 
                            (@BankID, @AccNo, @Type, @BankAccNo, @DefaultBank, @AllowOD, @ODAmt, @ODRate, @Chqformat)";

                        await conn.ExecuteAsync(insertBankAccSql, new
                        {
                            BankID = dto.BankID,
                            AccNo = dto.AccNo,
                            Type = string.IsNullOrWhiteSpace(dto.Type) ? "Current" : dto.Type,
                            BankAccNo = dto.BankAccNo,
                            DefaultBank = dto.DefaultBank ? 1 : 0,
                            AllowOD = dto.AllowOD ? 1 : 0,
                            ODAmt = dto.ODAmt,
                            ODRate = dto.ODRate,
                            Chqformat = dto.Chqformat
                        }, transaction: tran);
                    }
                    else
                    {
                        // 1. Update Accounts GL Title and Opening details
                        const string updateAccSql = @"
                            UPDATE Accounts 
                            SET AccTitle = @AccTitle,
                                openbal = @OpenBal,
                                opendate = @OpenDate
                            WHERE AccNo = @AccNo";

                        await conn.ExecuteAsync(updateAccSql, new
                        {
                            AccTitle = dto.AccTitle,
                            OpenBal = dto.OpenBal,
                            OpenDate = dto.OpenDate,
                            AccNo = dto.AccNo
                        }, transaction: tran);

                        // 2. Default bank exclusivity
                        if (dto.DefaultBank)
                        {
                            await conn.ExecuteAsync(
                                "UPDATE BankAccounts SET DefaultBank = 0 WHERE AccNo <> @AccNo",
                                new { dto.AccNo }, transaction: tran);
                        }

                        // 3. Update BankAccounts
                        const string updateBankAccSql = @"
                            UPDATE BankAccounts 
                            SET BankID = @BankID,
                                Type = @Type,
                                BankAccNo = @BankAccNo,
                                DefaultBank = @DefaultBank,
                                AllowOD = @AllowOD,
                                ODAmt = @ODAmt,
                                ODRate = @ODRate,
                                Chqformat = @Chqformat
                            WHERE AccNo = @AccNo";

                        await conn.ExecuteAsync(updateBankAccSql, new
                        {
                            BankID = dto.BankID,
                            Type = string.IsNullOrWhiteSpace(dto.Type) ? "Current" : dto.Type,
                            BankAccNo = dto.BankAccNo,
                            DefaultBank = dto.DefaultBank ? 1 : 0,
                            AllowOD = dto.AllowOD ? 1 : 0,
                            ODAmt = dto.ODAmt,
                            ODRate = dto.ODRate,
                            Chqformat = dto.Chqformat,
                            AccNo = dto.AccNo
                        }, transaction: tran);
                    }

                    tran.Commit();
                    return new BankOperationResult
                    {
                        Success = true,
                        Message = dto.IsNew ? "Bank account created successfully." : "Bank account updated successfully.",
                        GeneratedAccNo = dto.AccNo,
                        BankID = dto.BankID
                    };
                }
                catch (Exception ex)
                {
                    tran.Rollback();
                    return new BankOperationResult { Success = false, Message = $"Failed to save bank account: {ex.Message}" };
                }
            }
        }
    }

    public async Task<BankOperationResult> SetDefaultBankAccountAsync(string accNo)
    {
        using (var conn = new SqlConnection(_connectionString))
        {
            await conn.OpenAsync();
            using (var tran = conn.BeginTransaction())
            {
                try
                {
                    await conn.ExecuteAsync("UPDATE BankAccounts SET DefaultBank = 0", transaction: tran);
                    await conn.ExecuteAsync("UPDATE BankAccounts SET DefaultBank = 1 WHERE AccNo = @AccNo", new { AccNo = accNo }, transaction: tran);
                    tran.Commit();
                    return new BankOperationResult { Success = true, Message = "Default bank account updated successfully.", GeneratedAccNo = accNo };
                }
                catch (Exception ex)
                {
                    tran.Rollback();
                    return new BankOperationResult { Success = false, Message = $"Failed to set default bank account: {ex.Message}" };
                }
            }
        }
    }

    public async Task<BankOperationResult> UpdateOverdraftAsync(OverdraftConfigDto dto)
    {
        using (var conn = new SqlConnection(_connectionString))
        {
            try
            {
                const string sql = @"
                    UPDATE BankAccounts 
                    SET AllowOD = @AllowOD,
                        ODAmt = @ODAmt,
                        ODRate = @ODRate
                    WHERE AccNo = @AccNo";

                await conn.ExecuteAsync(sql, new
                {
                    AllowOD = dto.AllowOD ? 1 : 0,
                    ODAmt = dto.ODAmt,
                    ODRate = dto.ODRate,
                    AccNo = dto.AccNo
                });

                return new BankOperationResult { Success = true, Message = "Overdraft settings updated successfully." };
            }
            catch (Exception ex)
            {
                return new BankOperationResult { Success = false, Message = $"Failed to update overdraft: {ex.Message}" };
            }
        }
    }

    public async Task<BankOperationResult> DeleteBankAccountAsync(string accNo)
    {
        using (var conn = new SqlConnection(_connectionString))
        {
            await conn.OpenAsync();

            // 1. Guard against existing voucher / ledger transactions
            int voucherCount = await conn.ExecuteScalarAsync<int>(
                @"SELECT 
                    (SELECT COUNT(1) FROM Vouchers WHERE Accno = @AccNo) +
                    (SELECT COUNT(1) FROM VouchersForAuthorization WHERE Accno = @AccNo)",
                new { AccNo = accNo });

            if (voucherCount > 0)
            {
                return new BankOperationResult
                {
                    Success = false,
                    Message = $"Cannot delete this bank account because it has {voucherCount} posted voucher(s) in the General Ledger."
                };
            }

            // 2. Guard against existing issued cheques
            int issuedChqCount = await conn.ExecuteScalarAsync<int>(
                @"SELECT COUNT(1) 
                  FROM ChqList cl 
                  INNER JOIN ChqBooks cb ON cl.ChqBookNo = cb.ChqBookNo 
                  WHERE cb.AccNo = @AccNo AND cl.Issued = 1",
                new { AccNo = accNo });

            if (issuedChqCount > 0)
            {
                return new BankOperationResult
                {
                    Success = false,
                    Message = $"Cannot delete this bank account because {issuedChqCount} cheque(s) have already been issued from its cheque book(s)."
                };
            }

            // 3. Guard against non-zero balance
            decimal balance = await conn.ExecuteScalarAsync<decimal>(
                "SELECT ISNULL(Balance, 0) FROM Accounts WHERE AccNo = @AccNo", new { AccNo = accNo });

            if (balance != 0)
            {
                return new BankOperationResult
                {
                    Success = false,
                    Message = $"Cannot delete this bank account because it currently has a non-zero balance ({balance:N2})."
                };
            }

            // Perform deletion in atomic transaction in exact FK order
            using (var tran = conn.BeginTransaction())
            {
                try
                {
                    // Clean up unissued cheque list entries and cheque books
                    await conn.ExecuteAsync(@"
                        DELETE cl 
                        FROM ChqList cl
                        INNER JOIN ChqBooks cb ON cl.ChqBookNo = cb.ChqBookNo
                        WHERE cb.AccNo = @AccNo", new { AccNo = accNo }, transaction: tran);

                    await conn.ExecuteAsync(
                        "DELETE FROM ChqBooks WHERE AccNo = @AccNo", new { AccNo = accNo }, transaction: tran);

                    // Delete mapping from BankAccounts
                    await conn.ExecuteAsync(
                        "DELETE FROM BankAccounts WHERE AccNo = @AccNo", new { AccNo = accNo }, transaction: tran);

                    // Delete from Accounts master
                    await conn.ExecuteAsync(
                        "DELETE FROM Accounts WHERE AccNo = @AccNo", new { AccNo = accNo }, transaction: tran);

                    tran.Commit();
                    return new BankOperationResult { Success = true, Message = "Bank account deleted successfully." };
                }
                catch (Exception ex)
                {
                    tran.Rollback();
                    return new BankOperationResult { Success = false, Message = $"Failed to delete bank account: {ex.Message}" };
                }
            }
        }
    }

    public async Task<BankOperationResult> DeleteBankBranchAsync(int bankId)
    {
        using (var conn = new SqlConnection(_connectionString))
        {
            await conn.OpenAsync();

            int accountsCount = await conn.ExecuteScalarAsync<int>(
                "SELECT COUNT(1) FROM BankAccounts WHERE BankID = @BankID", new { BankID = bankId });

            if (accountsCount > 0)
            {
                return new BankOperationResult
                {
                    Success = false,
                    Message = $"Cannot delete this bank branch because {accountsCount} bank account(s) are associated with it. Delete or reassign those accounts first."
                };
            }

            try
            {
                await conn.ExecuteAsync("DELETE FROM BankList WHERE BankID = @BankID", new { BankID = bankId });
                return new BankOperationResult { Success = true, Message = "Bank branch deleted successfully." };
            }
            catch (Exception ex)
            {
                return new BankOperationResult { Success = false, Message = $"Failed to delete bank branch: {ex.Message}" };
            }
        }
    }

    // ─────────────────────────────────────────────────────────────
    // Legacy Compatibility Methods
    // ─────────────────────────────────────────────────────────────

    public async Task<List<BankListModel>> GetBankList()
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            var cond = " ORDER BY Bank, City ASC";
            string sql = @" SELECT AccNo AS BankAccNo, Fax as FaxNo, * FROM BankList " + cond;
            var banklist = await db.QueryAsync<BankListModel>(sql);
            return banklist.ToList();
        }
    }

    public async Task AddBankAsync(BankListModel bank)
    {
        var dto = new SaveBankBranchDto
        {
            Bank = bank.Bank,
            Branch = bank.Branch,
            City = bank.City,
            Phone1 = bank.Phone1,
            Phone2 = bank.Phone2,
            Fax = bank.FaxNo,
            Email = bank.Email,
            Manager = bank.Manager,
            Address = bank.Address,
            Remittance = bank.Remittance
        };
        await SaveBankBranchAsync(dto, "System", "Web");
    }

    public async Task UpdateBankAsync(BankListModel bank)
    {
        var dto = new SaveBankBranchDto
        {
            BankID = bank.BankID,
            Bank = bank.Bank,
            Branch = bank.Branch,
            City = bank.City,
            Phone1 = bank.Phone1,
            Phone2 = bank.Phone2,
            Fax = bank.FaxNo,
            Email = bank.Email,
            Manager = bank.Manager,
            Address = bank.Address,
            Remittance = bank.Remittance
        };
        await SaveBankBranchAsync(dto, "System", "Web");
    }

    public async Task SaveNewBank(BankListModel newbank)
    {
        var branchDto = new SaveBankBranchDto
        {
            Bank = newbank.Bank,
            Branch = newbank.Branch,
            City = newbank.City,
            Phone1 = newbank.Phone1,
            Phone2 = newbank.Phone2,
            Fax = newbank.FaxNo,
            Email = newbank.Email,
            Manager = newbank.Manager,
            Address = newbank.Address,
            Remittance = newbank.Remittance
        };

        var branchResult = await SaveBankBranchAsync(branchDto, "System", "Web");
        if (branchResult.Success && !string.IsNullOrWhiteSpace(newbank.BankAccNo))
        {
            var accDto = new SaveBankAccountDto
            {
                BankID = branchResult.BankID,
                AccTitle = $"{newbank.Bank} - {newbank.Branch}",
                BankAccNo = newbank.BankAccNo,
                Type = string.IsNullOrWhiteSpace(newbank.AccType) ? "Current" : newbank.AccType,
                IsNew = true
            };
            await SaveBankAccountAsync(accDto, "System", "Web");
        }
    }

    public async Task EditBankAsync(BankListModel newbank, int bankid)
    {
        var dto = new SaveBankBranchDto
        {
            BankID = bankid,
            Bank = newbank.Bank,
            Branch = newbank.Branch,
            City = newbank.City,
            Phone1 = newbank.Phone1,
            Phone2 = newbank.Phone2,
            Fax = newbank.FaxNo,
            Email = newbank.Email,
            Manager = newbank.Manager,
            Address = newbank.Address,
            Remittance = newbank.Remittance
        };
        await SaveBankBranchAsync(dto, "System", "Web");
    }

    public async Task DeleteBankAsync(int bankid)
    {
        await DeleteBankBranchAsync(bankid);
    }
}
