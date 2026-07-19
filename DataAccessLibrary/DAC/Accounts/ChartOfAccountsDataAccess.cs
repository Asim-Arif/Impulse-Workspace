using Dapper;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Configuration;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.IO;
using System.Reflection.PortableExecutable;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Accounts
{
    public class ChartOfAccountsDataAccess : IChartOfAccountsDataAccess
    {
        private readonly string _connectionString;
        public ChartOfAccountsDataAccess(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection");
        }
        public async Task<List<ChartOfAccountsModel>> GetChartOfAccountsList(bool showinactive, string filterAccNo = null)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    // Build WHERE / ORDER BY using only columns that exist in VAccounts
                    var cond = "";
                    if (showinactive == false)
                    {
                        cond = "WHERE A.Active = 1";
                    }
                    else
                    {
                        cond = "WHERE 1=1";
                    }

                    if (!string.IsNullOrEmpty(filterAccNo))
                    {
                        cond += " AND (A.AccNo = @filterAccNo OR A.SubAccOf LIKE @filterAccNo + '%')";
                    }

                    cond += " ORDER BY A.Type, A.AccNo";

                    // VAccounts = Accounts JOIN Heads ON Type=code, columns:
                    // MainHead, AccNo, AccTitle, SubAccOf, Type, openbal, opendate,
                    // Balance, Parent, Active, CreateDT, UserName, ComputerName, ParentAccTitle
                    string sql = @"
                    SELECT
                        A.AccNo,
                        A.AccTitle,
                        A.openbal       AS OpeningBalance,
                        A.Balance,
                        A.Parent        AS AccType,
                        ISNULL(G.GroupName, '') AS ReportingGroup,
                        ISNULL(AG.AccGroup_RefID, 0) AS AccGroup_RefID,
                        'No'          AS Budget,
                        'No'          AS Budget_G,
                        0             AS Taxpercentage,
                        A.Active,
                        A.Type          AS HeadTypeNo,
                        A.MainHead      AS HeadType,
                        A.SubAccOf,
                        ISNULL(A.ParentAccTitle,'') AS SubAccTitle,
                        A.opendate      AS OpenDate,
                        CAST(1 AS BIT) AS bMainhead
                    FROM VAccounts A
                    LEFT JOIN AccountGroupsForExpenseReport AG ON A.AccNo = AG.AccNo
                    LEFT JOIN GroupsForExpenseReport G ON AG.AccGroup_RefID = G.EntryID " + cond;

                    var chartOfAccountsList = await db.QueryAsync<ChartOfAccountsModel>(sql, new { filterAccNo });

                    //Console.WriteLine($"Error fetching ChartOfAccounts: {sql}");

                    return chartOfAccountsList.ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching ChartOfAccounts: {ex.Message}");
                throw; // Rethrow or handle appropriately
            }
        }

        public async Task<List<ChartOfAccountsModel>> GetAccountsHeads()
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {

                    string sql = @"SELECT Head AS HeadType,Code AS HeadTypeNo FROM Heads ORDER BY Heads.Code ASC";

                    var chartOfAccountsList = await db.QueryAsync<ChartOfAccountsModel>(sql, new { });
                    return chartOfAccountsList.ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching ChartOfAccounts: {ex.Message}");
                throw; // Rethrow or handle appropriately
            }
        }

        public async Task<List<ChartOfAccountsModel>> GetSubAccOfAccounts(string acctype)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {

                    string sql = @"Select AccTitle AS SubAccTitle,AccNo As SubAccOf From Accounts Where Type=@acctype and Parent=1";
                    var chartOfAccountsList = await db.QueryAsync<ChartOfAccountsModel>(sql, new { acctype = acctype });
                    return chartOfAccountsList.ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching ChartOfAccounts: {ex.Message}");
                throw; // Rethrow or handle appropriately
            }
        }
        public async Task<string> GetNextAccountNumberAsync(string acctype, string accountname, Boolean isparent, string subaccof)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                // DB function: GetNextAccno(@AccountName, @AccType, @ParentAccount, @IsParent)
                const string query = "SELECT dbo.GetNextAccno(@AccountName, @AccType, @ParentAccount, @IsParent)";
                return await connection.QuerySingleOrDefaultAsync<string>(query, new
                {
                    AccountName = accountname,
                    AccType = acctype,
                    ParentAccount = subaccof,
                    IsParent = isparent
                });
            }
        }
        public async Task SaveNewAccount(ChartOfAccountsModel newaccount)
        {
            if (string.IsNullOrWhiteSpace(newaccount.AccTitle))
            {
                throw new Exception("Cannot proceed without account title.");
            }

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                try
                {
                    if (newaccount.BalType == 1)
                    {
                        newaccount.Balance = -newaccount.Balance;
                    }

                    string sql = @"
                    INSERT INTO Accounts 
                    (AccNo, AccTitle, SubAccOf, Type, openbal, opendate, Balance, Parent, Active, UserName, ComputerName) 
                    VALUES 
                    (@AccNo, @AccTitle, @SubAccOf, @HeadTypeNo, @OpenBal, @OpenDate, @Balance, @Parent, @Active, @UserName, @machineIP)";

                    var parameters = new
                    {
                        AccNo = newaccount.AccNo,
                        AccTitle = newaccount.AccTitle,
                        SubAccOf = newaccount.SubAccOf,
                        HeadTypeNo = newaccount.HeadTypeNo,
                        OpenBal = newaccount.Balance,
                        OpenDate = newaccount.OpenDate,
                        Balance = newaccount.Balance,
                        Active = 1,
                        Parent = newaccount.AccType,
                        UserName = "Developer",
                        machineIP = "Developer"
                    };

                    await db.ExecuteAsync(sql, parameters);
                }
                catch (Exception ex)
                {
                    throw new Exception("Error saving account: " + ex.Message);
                }
            }
        }
        public async Task EditAccount(ChartOfAccountsModel newaccount, string AccNo)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                try
                {
                    string sql = @"
                    UPDATE Accounts SET AccTitle=@AccTitle, openbal=@OpenBal, opendate=@OpenDate WHERE AccNo=@AccNo";

                    var parameters = new
                    {
                        AccNo = newaccount.AccNo,
                        AccTitle = newaccount.AccTitle,
                        OpenBal = newaccount.OpeningBalance,
                        OpenDate = newaccount.OpenDate
                    };
                    await db.ExecuteAsync(sql, parameters);
                }
                catch (Exception ex)
                {
                    throw new Exception("Error Update account: " + ex.Message);
                }
            }
        }

        public async Task DeleteAccountAsync(string AccNo)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                try
                {
                    string sql = @"
                    DELETE FROM Accounts WHERE AccNo=@AccNo";
                    await db.ExecuteAsync(sql, new { AccNo });
                }
                catch (Exception ex)
                {
                    throw new Exception("Error Delete Account: " + ex.Message);
                }
            }
        }

        public async Task UpdateAccountStatusAsync(string AccNo, Boolean bactive)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                try
                {
                    string sql = "";
                    if (bactive == true)
                    {
                        sql = @"Update Accounts SET Active=0 WHERE AccNo=@AccNo";
                    }
                    else
                    {
                        sql = @"Update Accounts SET Active=1 WHERE AccNo=@AccNo";
                    }
                    await db.ExecuteAsync(sql, new { AccNo });
                }
                catch (Exception ex)
                {
                    throw new Exception("Error Updating Account: " + ex.Message);
                }
            }
        }

        public async Task<List<CashFlowHeadsModel>> GetCashFlowHeads()
        {
            // Table 'CashFlowHeadsForStatement' does not exist in this database schema.
            // Return empty list until the table is created.
            await Task.CompletedTask;
            return new List<CashFlowHeadsModel>();
        }

        public async Task UpdateCashFlowHeads(ChartOfAccountsModel entry, string AccNo, long CFHFS_RefID = 0, int itype = 0)
        {
            // Columns Marked_CF, CFHFS_RefID, Marked_CF_Nve do not exist on the Accounts table
            // in this database schema. This method is a no-op until they are added.
            await Task.CompletedTask;
        }

        public async Task DeleteAccountGroupsForBalancesheet(string AccNo)
        {
            // Table 'AccountGroupsForBalancesheet' does not exist in this database schema.
            // This method is a no-op until the table is created.
            await Task.CompletedTask;
        }

        public async Task SaveAccountGroupsForBalancesheet(ChartOfAccountsModel newaccount, string accno, long AccGroup_RefID = 0, int itype = 0)
        {
            // Table 'AccountGroupsForBalancesheet' and columns 'Marked', 'CFHFS_RefID', 'Marked_CF_Nve'
            // do not exist in this database schema. This method is a no-op until they are created.
            await Task.CompletedTask;
        }
        public async Task<List<NotesToAccountsModel>> GetNotesToAccountsHeads()
        {
            // Table 'GroupsForFinancialReport_Sub' does not exist in this database schema.
            // Return empty list until the table is created.
            await Task.CompletedTask;
            return new List<NotesToAccountsModel>();
        }

        public async Task<List<ReportingGroupModel>> GetReportingGroupsAsync()
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = @"SELECT EntryID, GroupName, Type, SortNo FROM GroupsForExpenseReport ORDER BY SortNo ASC";
                    var groups = await db.QueryAsync<ReportingGroupModel>(sql);
                    return groups.ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching reporting groups: {ex.Message}");
                throw;
            }
        }

        public async Task AddReportingGroupAsync(ReportingGroupModel model)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = "INSERT INTO GroupsForExpenseReport (GroupName, Type, SortNo) VALUES (@GroupName, @Type, @SortNo)";
                await db.ExecuteAsync(sql, new { model.GroupName, model.Type, model.SortNo });
            }
        }

        public async Task UpdateReportingGroupDetailsAsync(ReportingGroupModel model)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = "UPDATE GroupsForExpenseReport SET GroupName = @GroupName, Type = @Type, SortNo = @SortNo WHERE EntryID = @EntryID";
                await db.ExecuteAsync(sql, new { model.GroupName, model.Type, model.SortNo, model.EntryID });
            }
        }

        public async Task DeleteReportingGroupAsync(long entryId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = "DELETE FROM GroupsForExpenseReport WHERE EntryID = @EntryID";
                await db.ExecuteAsync(sql, new { EntryID = entryId });
            }
        }

        public async Task UpdateReportingGroupAsync(string AccNo, long AccGroup_RefID)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                try
                {
                    if (AccGroup_RefID == 0)
                    {
                        // Remove mapping
                        string deleteSql = "DELETE FROM AccountGroupsForExpenseReport WHERE AccNo = @AccNo";
                        await db.ExecuteAsync(deleteSql, new { AccNo });
                    }
                    else
                    {
                        // Update or Insert mapping
                        string checkSql = "SELECT COUNT(1) FROM AccountGroupsForExpenseReport WHERE AccNo = @AccNo";
                        int exists = await db.ExecuteScalarAsync<int>(checkSql, new { AccNo });

                        if (exists > 0)
                        {
                            string updateSql = "UPDATE AccountGroupsForExpenseReport SET AccGroup_RefID = @AccGroup_RefID WHERE AccNo = @AccNo";
                            await db.ExecuteAsync(updateSql, new { AccGroup_RefID, AccNo });
                        }
                        else
                        {
                            string insertSql = "INSERT INTO AccountGroupsForExpenseReport (AccGroup_RefID, AccNo) VALUES (@AccGroup_RefID, @AccNo)";
                            await db.ExecuteAsync(insertSql, new { AccGroup_RefID, AccNo });
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Error updating reporting group: " + ex.Message);
                }
            }
        }

        public async Task<List<ChartOfAccountsModel>> GetParentCategoriesAsync()
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = @"SELECT AccNo, AccTitle FROM Accounts WHERE Parent=1 AND Active=1 ORDER BY AccNo";
                    var chartOfAccountsList = await db.QueryAsync<ChartOfAccountsModel>(sql);
                    return chartOfAccountsList.ToList();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error fetching Parent Categories: " + ex.Message);
            }
        }

        public async Task<List<ChartOfAccountsModel>> GetChildAccountsAsync(string parentAccNo)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = @"SELECT AccNo, SubAccOf, AccTitle FROM Accounts WHERE SubAccOf = @parentAccNo ORDER BY AccNo";
                    var chartOfAccountsList = await db.QueryAsync<ChartOfAccountsModel>(sql, new { parentAccNo });
                    return chartOfAccountsList.ToList();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error fetching Child Accounts: " + ex.Message);
            }
        }

        public async Task ChangeAccountCategoryAsync(string oldAccNo, string newAccNo, string newParentAccNo, string userName, string machineName)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();
                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        string typeCode = newParentAccNo.Length >= 2 ? newParentAccNo.Substring(0, 2) : newParentAccNo;

                        string insertAudit = @"INSERT INTO AccountHeadChanges(OriginalAccNo, NewAccNo, UserName, MachineName) 
                                             VALUES(@oldAccNo, @newAccNo, @userName, @machineName)";
                        await connection.ExecuteAsync(insertAudit, new { oldAccNo, newAccNo, userName, machineName }, transaction);

                        string updateAccounts = @"UPDATE Accounts SET Type=@typeCode, SubAccOf=@newParentAccNo, AccNo=@newAccNo WHERE AccNo=@oldAccNo";
                        await connection.ExecuteAsync(updateAccounts, new { typeCode, newParentAccNo, newAccNo, oldAccNo }, transaction);

                        // Update dependent tables
                        var updateQueries = new List<string>
                        {
                            "UPDATE Makers SET AccNo=@newAccNo WHERE AccNo=@oldAccNo",
                            "UPDATE ForeignCustomers SET AccNo=@newAccNo WHERE AccNo=@oldAccNo",
                            "UPDATE VouchersDeleted SET AccNo=@newAccNo WHERE AccNo=@oldAccNo",
                            "UPDATE Departments SET AccruedAccNo=@newAccNo WHERE AccruedAccNo=@oldAccNo",
                            "UPDATE Departments SET SalaryAccNo=@newAccNo WHERE SalaryAccNo=@oldAccNo",
                            "UPDATE Departments SET AccruedFoodAccNo=@newAccNo WHERE AccruedFoodAccNo=@oldAccNo",
                            "UPDATE Departments SET AccruedUnionAccNo=@newAccNo WHERE AccruedUnionAccNo=@oldAccNo",
                            "UPDATE Departments SET SchAccNo=@newAccNo WHERE SchAccNo=@oldAccNo",
                            "UPDATE Departments SET AccruedEOBIAccNo=@newAccNo WHERE AccruedEOBIAccNo=@oldAccNo",
                            "UPDATE Departments SET TaxAccNo=@newAccNo WHERE TaxAccNo=@oldAccNo",
                            "UPDATE Cheque SET AccNo=@newAccNo WHERE AccNo=@oldAccNo",
                            "UPDATE VendOrders SET VendID=@newAccNo WHERE VendID=@oldAccNo",
                            "UPDATE VendRcvd SET VendID=@newAccNo WHERE VendID=@oldAccNo"
                        };

                        foreach (var sql in updateQueries)
                        {
                            await connection.ExecuteAsync(sql, new { newAccNo, oldAccNo }, transaction);
                        }

                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        throw new Exception("Error changing Account Category: " + ex.Message);
                    }
                }
            }
        }

        public async Task ReindexFinancialBalancesAsync()
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();
                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        // 1. Recalculate all Voucher balances using a high-performance window function
                        string updateVouchersSql = @"
                            WITH CTE AS (
                                SELECT 
                                    v.AccNo, v.Sno, v.Debit, v.Credit, v.Balance,
                                    COALESCE(a.OpenBal, 0) + SUM(ISNULL(v.Debit, 0) - ISNULL(v.Credit, 0)) OVER (PARTITION BY v.AccNo ORDER BY v.Sno ROWS UNBOUNDED PRECEDING) AS CalculatedBalance
                                FROM Vouchers v
                                INNER JOIN Accounts a ON v.AccNo = a.AccNo
                            )
                            UPDATE CTE SET Balance = CalculatedBalance;";
                        
                        // Execute timeout increased in case there are millions of vouchers
                        await connection.ExecuteAsync(updateVouchersSql, null, transaction, commandTimeout: 300);

                        // 2. Update Accounts table with the final calculated balance
                        string updateAccountsSql = @"
                            UPDATE Accounts
                            SET Balance = COALESCE(
                                (SELECT TOP 1 v.Balance 
                                 FROM Vouchers v 
                                 WHERE v.AccNo = Accounts.AccNo 
                                 ORDER BY v.Sno DESC), 
                                Accounts.OpenBal);";

                        await connection.ExecuteAsync(updateAccountsSql, null, transaction, commandTimeout: 300);

                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        throw new Exception("Error during financial re-indexing: " + ex.Message);
                    }
                }
            }
        }
    }
}


