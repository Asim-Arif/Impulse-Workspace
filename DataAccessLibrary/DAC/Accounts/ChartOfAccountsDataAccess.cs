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
        public async Task<List<ChartOfAccountsModel>> GetChartOfAccountsList(bool showinactive)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    var cond = "";
                    if (showinactive == false)
                    {
                        cond = "WHERE Active = 1 ORDER BY code, DisplayLine, SortOrder";
                    }
                    else {
                        cond = " ORDER BY code, DisplayLine, SortOrder";
                    }


                    string sql = @"
                    SELECT 
                    AccNo,AccTitle,OpenBal AS OpeningBalance,Balance,Parent As AccType,'' AS ReportingGroup,
                    CASE WHEN BudgetAccount=1 THEN 'Yes' ELSE 'No' END AS Budget,
                    CASE WHEN BudgetAccount_Glance=1 THEN 'Yes' ELSE 'No' END AS Budget_G,
                    ISNULL(Tax_Per,0) AS Taxpercentage,Active,Type as HeadTypeNo,MainHead AS HeadType,
                    SubAccOf,'' AS SubAccTitle,OpenDate,SortOrder AS bMainhead
                    FROM VAccounts_Sorted " + cond;

                    var chartOfAccountsList = await db.QueryAsync<ChartOfAccountsModel>(sql, new { });

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
                const string query = "SELECT dbo.GetNextAccountNo(@acctype,@accountname,@isparent,@subaccof)";
                return await connection.QuerySingleOrDefaultAsync<string>(query, new { acctype, accountname, isparent, subaccof });
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
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = @"SELECT EntryID,Description FROM CashFlowHeadsForStatement T1 ORDER BY T1.Description ASC";
                    var heads = await db.QueryAsync<CashFlowHeadsModel>(sql, new { });
                    return heads.ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching ChartOfAccounts: {ex.Message}");
                throw;
            }
        }

        public async Task UpdateCashFlowHeads(ChartOfAccountsModel entry, string AccNo, long CFHFS_RefID = 0, int itype = 0)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                try
                {
                    bool bMarked_CF_Nve = entry.Marked_CF_Nve;
                    string sql = "";
                    if (itype == 1)
                    {
                        sql = @"Update Accounts SET Marked_CF=0 WHERE AccNo=@AccNo";
                    }
                    else if (itype == 2)
                    {
                        sql = @"Update Accounts SET Marked_CF=1 WHERE AccNo=@AccNo";
                    }
                    else
                    {
                        sql = @"Update Accounts SET CFHFS_RefID=@CFHFS_RefID,Marked_CF_Nve=@bMarked_CF_Nve WHERE AccNo=@AccNo";
                    }
                    await db.ExecuteAsync(sql, new { AccNo, CFHFS_RefID, bMarked_CF_Nve });
                }
                catch (Exception ex)
                {
                    throw new Exception("Error Updating Account: " + ex.Message);
                }
            }
        }

        public async Task DeleteAccountGroupsForBalancesheet(string AccNo)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                try
                {
                    string sql = @"
                    DELETE FROM AccountGroupsForBalancesheet WHERE Left(Accno,2)=@AccNo";
                    await db.ExecuteAsync(sql, new { AccNo });
                }
                catch (Exception ex)
                {
                    throw new Exception("Error Delete Account: " + ex.Message);
                }
            }
        }

        public async Task SaveAccountGroupsForBalancesheet(ChartOfAccountsModel newaccount, string accno, long AccGroup_RefID = 0, int itype = 0)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                try
                {
                    string sql = "";
                    if (itype == 1)
                    {
                        sql = @"
                        INSERT INTO AccountGroupsForBalancesheet 
                        (AccGroup_RefID,AccNo) 
                        SELECT @AccGroup_RefID,Accno FROM Accounts WHERE Left(Accno,2)=@accno AND Parent=0";
                    }
                    else if (itype == 2)
                    {
                        sql = @"
                        INSERT INTO AccountGroupsForBalancesheet 
                        (AccGroup_RefID,AccNo,Marked) 
                        VALUES
                        (@AccGroup_RefID,@accno,@Marked)  ";
                    }
                    else if (itype == 3)    ////// -ve marking Status
                    {
                        sql = @"UPDATE AccountGroupsForBalanceSheet SET Marked=0 WHERE AccNo=@AccNo";
                    }
                    else if (itype == 4)   ////// -ve marking Status
                    {
                        sql = @"UPDATE AccountGroupsForBalanceSheet SET Marked=1 WHERE AccNo=@AccNo";
                    }
                    else if (itype == 5)   ////// TB Status
                    {
                        sql = @"UPDATE Accounts SET Marked=0 WHERE AccNo=@AccNo";
                    }
                    else if (itype == 6)   ////// TB Status
                    {
                        sql = @"UPDATE Accounts SET Marked=1 WHERE AccNo=@AccNo";
                    }

                    else if (itype == 7)   ////// Break Notes Link
                    {
                        sql = @"
                                DELETE FROM AccountGroupsForBalancesheet WHERE AccNo=@AccNo
                                UPDATE Accounts SET Marked=0 WHERE AccNo=@AccNo";
                    }
                    else if (itype == 8)   ////// Break Cash Flow Link
                    {
                        sql = @"UPDATE Accounts SET CFHFS_RefID=0,Marked_CF_Nve=0 WHERE AccNo=@AccNo";
                    }
                    else
                    {
                        sql = @"
                        INSERT INTO AccountGroupsForBalancesheet 
                        (AccGroup_RefID,AccNo,Marked) 
                        VALUES
                        (@AccGroup_RefID,@accno,@Marked)  ";
                    }
                    var parameters = new
                    {
                        AccNo = accno,
                        AccGroup_RefID = AccGroup_RefID,
                        Marked = newaccount.Marked
                    };

                    await db.ExecuteAsync(sql, parameters);
                }
                catch (Exception ex)
                {
                    throw new Exception("Error saving account: " + ex.Message);
                }
            }
        }
        public async Task<List<NotesToAccountsModel>> GetNotesToAccountsHeads()
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = @"SELECT EntryID,GroupName FROM GroupsForFinancialReport_Sub T1 ORDER BY T1.Type ASC";
                    var heads = await db.QueryAsync<NotesToAccountsModel>(sql, new { });
                    return heads.ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error on fetching: {ex.Message}");
                throw;
            }
        }

    }
}


