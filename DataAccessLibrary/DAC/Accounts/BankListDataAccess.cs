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
using System.Linq.Expressions;
using System.Reflection.PortableExecutable;
using System.Threading.Tasks;

public class BankListDataAccess : IBankListDataAccess
{
    private readonly string _connectionString;

    public BankListDataAccess(IConfiguration config)
    {
        _connectionString = config.GetConnectionString("DefaultConnection");
    }

    public async Task<List<BankListModel>> GetBankList()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var cond= " ORDER BY Bank,City ASC";
                string sql = @" SELECT AccNo AS BankAccNo,Fax as FaxNo,* FROM BankList " + cond;
                var banklist = await db.QueryAsync<BankListModel>(sql, new { });
                return banklist.ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching ChartOfAccounts: {ex.Message}");
            throw; 
        }
    }

    public async Task AddBankAsync(BankListModel bank)
    {
        using var conn = new SqlConnection(_connectionString);
        await conn.ExecuteAsync(
            "INSERT INTO BankProfiles (BankName, Branch, AccountNumber, IBAN, SwiftCode, Address) " +
            "VALUES (@BankName, @Branch, @AccountNumber, @IBAN, @SwiftCode, @Address)", bank);
    }

    public async Task UpdateBankAsync(BankListModel bank)
    {
        using var conn = new SqlConnection(_connectionString);
        await conn.ExecuteAsync(
            "UPDATE BankProfiles SET BankName=@BankName, Branch=@Branch, AccountNumber=@AccountNumber, " +
            "IBAN=@IBAN, SwiftCode=@SwiftCode, Address=@Address WHERE BankID=@BankID", bank);
    }

    public async Task SaveNewBank(BankListModel newbank)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"
                    INSERT INTO BankList 
                    (Bank,Branch,City,Phone1,Phone2,Fax,Email,Manager,Address,AccNo,ShowInCIP,COA,Remittance) 
                    OUTPUT INSERTED.BankID
                    VALUES 
                    (@Bank,@Branch,@City,@Phone1,@Phone2,@FaxNo,@Email,@Manager,@Address,@BankAccNo,@ShowInCIP,@COA,@Remittance)";

                var parameters = new
                {
                    Bank = newbank.Bank,
                    Branch = newbank.Branch,
                    City = newbank.City,
                    Phone1 = newbank.Phone1,
                    Phone2 = newbank.Phone2,
                    FaxNo = newbank.FaxNo,
                    Email = newbank.Email,
                    Manager = newbank.Manager,
                    Address = newbank.Address,
                    BankAccNo = newbank.BankAccNo,
                    ShowInCIP = newbank.ShowInCIP,
                    COA = newbank.COA,
                    Remittance=newbank.Remittance
                };
                //await db.ExecuteAsync(sql, parameters);
                int newBankId = await db.ExecuteScalarAsync<int>(sql, parameters);

                /////////////////////////////// IF COA MARKED THEN ADD IN ACCOUNTS
                if (newbank.COA == true) 
                {
                    string sql_bankacc = @"INSERT INTO BankAccounts(BankID,AccNo,Type,BankAccNo) 
                    VALUES (@BankID, @AccNo, @AccType, @BankAccNo)";
                    var parameters_bankacc = new
                    {
                        BankID = newBankId,
                        AccNo = newbank.AccNo,
                        AccType = newbank.AccType,
                        BankAccNo = newbank.BankAccNo,
                    };
                    await db.ExecuteAsync(sql_bankacc, parameters_bankacc);

                    string sql_acc = @"INSERT INTO Accounts (AccNo, AccTitle, SubAccOf,Type, openbal, opendate,Balance) 
                    VALUES (@AccNo, @AccTitle, @SubAccof,@AccType, @OpenBal, @OpenDate,@Balance)";
                    var parameters_acc = new
                    {
                        AccNo = newbank.AccNo,
                        AccTitle = newbank.Bank,
                        SubAccof= newbank.SubAccOf,
                        AccType = newbank.AccType,
                        OpenBal = 0.00,
                        Balance = 0.00,
                        OpenDate = DateTime.Now.Date
                    };
                    await db.ExecuteAsync(sql_acc, parameters_acc);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error saving account: " + ex.Message);
                Console.WriteLine($"Error fetching Data: {ex.Message}");

            }
        }
    }

    public async Task EditBankAsync(BankListModel newbank, int bankid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql =
                    @"UPDATE BankList SET Bank=@Bank,Branch=@Branch,City=@City,
                    Phone1=@Phone1,Phone2=@Phone2,Fax=@Fax,Email=@Email,Manager=@Manager,
                    Address=@Address,AccNo=@AccNo,Remittance=@Remittance,ShowInCIP=@ShowInCIP,COA=@COA 
                    WHERE BankID=@bankid";

                var parameters = new
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
                    AccNo = newbank.BankAccNo,
                    Remittance = newbank.Remittance,
                    ShowInCIP = newbank.ShowInCIP,
                    COA = newbank.COA,
                    bankid = bankid
                };
                await db.ExecuteAsync(sql, parameters);
            }
            catch (Exception ex)
            {
                throw new Exception("Error Update account: " + ex.Message);
            }
        }
    }

    public async Task DeleteBankAsync(int bankid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"
                    DELETE FROM Accounts WHERE AccNo=@bankid";
                await db.ExecuteAsync(sql, new { bankid });
            }
            catch (Exception ex)
            {
                throw new Exception("Error Delete Account: " + ex.Message);
            }
        }
    }



}
