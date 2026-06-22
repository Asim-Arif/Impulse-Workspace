using Dapper;
using DataAccessLibrary;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
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

public class FinancialStatementDataAccess : IFinancialStatementDataAccess
{
    private readonly string _connectionString;
    private readonly ILogger<FinancialStatementDataAccess> _logger;

    private readonly IDBHelper _db;

    public FinancialStatementDataAccess(IConfiguration config, IDBHelper db)
    {
        _connectionString = config.GetConnectionString("DefaultConnection");
        _db = db;
    }

    public async Task<List<MainHeadsModel>> GetMainHeadsList()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var cond= " ORDER BY SortNo ASC";
                string sql = @" SELECT EntryID,GroupName,SortNo,SequenceNo,Description FROM GroupsForFinancialReport " + cond;
                var entrieslist = await db.QueryAsync<MainHeadsModel>(sql, new { });
                return entrieslist.ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching: {ex.Message}");
            throw; 
        }
    }

    public async Task<List<SubHeadsModel>> GetSubHeadList()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var cond = " ORDER BY ENTRYID ASC";
                string sql = @" SELECT EntryID,GroupName_SUb AS GroupName,GroupName AS MainGroupName,SortNo AS SortNo,Type,SubDesc,Notes,ShowDetail,bShowSubHeadDetail_NTA FROM VFinancialGroups  " + cond;
                var entrieslist = await db.QueryAsync<SubHeadsModel>(sql, new { });
                return entrieslist.ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching : {ex.Message}");
            throw;
        }
    }
    public async Task<List<GroupsAccountsModel>> GetGroupsAccountsList()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var cond = " ORDER BY Type ASC,MainGroupName ASC";
                string sql = @" SELECT EntryID,GroupName,SortNo,Type,MainGroupName FROM GroupsForExpenseReport " + cond;
                var entrieslist = await db.QueryAsync<GroupsAccountsModel>(sql, new { });
                return entrieslist.ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching : {ex.Message}");
            throw;
        }
    }

    public async Task<bool> SaveMainHead(MainHeadsModel mainheads)
    {
        if (string.IsNullOrWhiteSpace(mainheads.GroupName))
        {
            throw new Exception("Cannot Proceed Without Group Name.");
        }
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"
                    INSERT INTO GroupsForFinancialReport 
                    (GroupName, SortNo, SequenceNo, Description)
                    VALUES 
                    (@GroupName, @SortNo, @SequenceNo, @Description)";
                var parameters = new
                {
                    GroupName = mainheads.GroupName,
                    SortNo = mainheads.SortNo,
                    SequenceNo = mainheads.SequenceNo,
                    Description = mainheads.Description,
                };
                int isavedrecords= await db.ExecuteAsync(sql, parameters);
                return isavedrecords > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Save: " + ex.Message);
            }
        }
    }

    public async Task<bool> SaveSubHead(SubHeadsModel subheads)
    {
        if (string.IsNullOrWhiteSpace(subheads.GroupName))
        {
            throw new Exception("Cannot Proceed Without Group Name.");
        }
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"
                    INSERT INTO GroupsForFinancialReport_Sub
                    (GroupName,Type,SubDesc,Notes,ShowDetail,bShowSubHeadDetail_NTA)
                    VALUES 
                    (@GroupName, @Type, @SubDesc, @Notes,@ShowDetail,@bShowSubHeadDetail_NTA)";
                var parameters = new
                {
                    GroupName = subheads.GroupName,
                    Type = subheads.Type,
                    SubDesc = subheads.SubDesc,
                    Notes = subheads.Notes,
                    ShowDetail = subheads.ShowDetail,
                    bShowSubHeadDetail_NTA = subheads.bShowSubHeadDetail_NTA
                };
                int isavedrecords = await db.ExecuteAsync(sql, parameters);
                return isavedrecords > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Save: " + ex.Message);
            }
        }
    }
    public async Task<bool> DeleteMainHead(long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"DELETE FROM GroupsForFinancialReport WHERE EntryID=@entryid";
                int rowsAffected = await db.ExecuteAsync(sql, new { entryid });
                return rowsAffected>0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Delete: " + ex.Message);
            }
        }
    }
    public async Task<bool> DeleteSubHead(long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"DELETE FROM GroupsForFinancialReport_Sub WHERE EntryID=@entryid";
                int rowsAffected = await db.ExecuteAsync(sql, new { entryid });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Delete: " + ex.Message);
            }
        }
    }
    public async Task<bool> EditMainHead(MainHeadsModel mainhead, long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql =
                    @"UPDATE GroupsForFinancialReport SET GroupName=@GroupName,SortNo=@SortNo,
                    Description=@Description
                    WHERE EntryID=@entryid";

                var parameters = new
                {
                    entryid= @entryid,
                    GroupName = mainhead.GroupName,
                    SortNo = mainhead.SortNo,
                    Description = mainhead.Description
                };
                int records= await db.ExecuteAsync(sql, parameters);
                return records> 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Update: " + ex.Message);
            }
        }
    }
    public async Task<bool> EditSubHead(SubHeadsModel subhead, long entryid,int itype)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = "";
                int records = 0;
                if (itype == 0)
                {
                    sql =
                        @"UPDATE GroupsForFinancialReport_Sub SET 
                        GroupName=@GroupName,Notes=@Notes,
                        SubDesc=@SubDesc
                        WHERE EntryID=@entryid";

                    var parameters = new
                    {
                        entryid = @entryid,
                        GroupName = subhead.GroupName,
                        SubDesc = subhead.SubDesc,
                        Notes = subhead.Notes,
                        bShowSubHeadDetail_NTA = subhead.bShowSubHeadDetail_NTA
                    };
                    records = await db.ExecuteAsync(sql, parameters);
                    return records > 0;
                }
                else 
                {
                    if (itype == 1)
                    {
                        sql =@"UPDATE GroupsForFinancialReport_Sub SET 
                                bShowSubHeadDetail_NTA=1 WHERE EntryID=@entryid";
                    }
                    else 
                    {
                        sql = @"UPDATE GroupsForFinancialReport_Sub SET 
                                bShowSubHeadDetail_NTA=0 WHERE EntryID=@entryid";
                    }
                    records = await db.ExecuteAsync(sql, new { entryid });
                    return records > 0;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Update: " + ex.Message);
            }
        }
    }

    public async Task<List<SubHeadsModel>> GetMainHeads()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"SELECT EntryID AS Type,GroupName MainGroupName FROM GroupsForFinancialReport";
                sql += " ORDER BY EntryID";
                return (await db.QueryAsync<SubHeadsModel>(sql)).ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching: {ex.Message}");
            throw; 
        }
    }


    public async Task<bool> DeleteSubHeadAccount(long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"DELETE FROM AccountGroup_HeadsDetail WHERE EntryID=@entryid";
                int rowsAffected = await db.ExecuteAsync(sql, new { entryid });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Delete: " + ex.Message);
            }
        }
    }
    public async Task<List<SubHeadAccountDetailModel>> GetAccountGroup_HeadsDetail(long AccGroup_RefID)
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var cond = "";
                cond=cond + "WHERE AccGroup_RefID=@AccGroup_RefID ORDER BY ENTRYID ASC";
                string sql = @"
                    SELECT EntryID,AccGroup_RefID,GroupName,AccNo,SubAccOf,NoteNo,Description,Notes,AccTitle,Marked FROM VAccountGroup_HeadsDetail " + cond;

                var entrieslist = await db.QueryAsync<SubHeadAccountDetailModel>(sql, new { AccGroup_RefID });
                return entrieslist.ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching : {ex.Message}");
            throw;
        }
    }

    public async Task<bool> SaveSubHeadDetail(SubHeadAccountDetailModel subheads)
    {
        if (string.IsNullOrWhiteSpace(subheads.Description))
        {
            throw new Exception("Cannot Proceed Without Group Name.");
        }
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"
                    INSERT INTO AccountGroup_HeadsDetail
                    (AccGroup_RefID,AccNo,NoteNo,Description,UserName,MachineName)
                    VALUES 
                    (@AccGroup_RefID, @AccNo, @NoteNo, @Description,@UserName,@MachineName)";
                var parameters = new
                {
                    AccGroup_RefID = subheads.AccGroup_RefID,
                    AccNo = subheads.AccNo,
                    NoteNo = subheads.NoteNo,
                    Description = subheads.Description,
                    UserName = subheads.UserName,
                    MachineName = subheads.MachineName
                };
                int isavedrecords = await db.ExecuteAsync(sql, parameters);
                return isavedrecords > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Save: " + ex.Message);
            }
        }
    }
    public async Task<bool> DeleteSubHeadEntry(long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"DELETE FROM AccountGroup_HeadsDetail WHERE EntryID=@entryid";
                int rowsAffected = await db.ExecuteAsync(sql, new { entryid });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Delete: " + ex.Message);
            }
        }
    }
    public async Task<bool> EditSubHeadDetailEntry(SubHeadAccountDetailModel mainhead, long entryid, int itype)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = "";
                int records = 0;
                if (itype == 0)
                {
                    sql =
                        @"UPDATE AccountGroup_HeadsDetail SET Description=@Description,NoteNo=@NoteNo
                        WHERE EntryID=@entryid";

                    var parameters = new
                    {
                        entryid = @entryid,
                        Description = mainhead.Description,
                        NoteNo = mainhead.NoteNo,
                    };
                    records = await db.ExecuteAsync(sql, parameters);
                }
                else
                {
                    if (itype == 1) 
                    { 
                        sql =
                        @"UPDATE AccountGroup_HeadsDetail SET Marked=0
                        WHERE EntryID=@entryid";
                    }
                    else
                    {
                        sql =
                        @"UPDATE AccountGroup_HeadsDetail SET Marked=1
                        WHERE EntryID=@entryid";
                    }
                        records = await db.ExecuteAsync(sql, new { entryid });
                }
                return records > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Update: " + ex.Message);
            }
        }
    }

    public async Task<bool> SaveGroupsAccounts(GroupsAccountsModel mainheads)
    {
        if (string.IsNullOrWhiteSpace(mainheads.GroupName))
        {
            throw new Exception("Cannot Proceed Without Group Name.");
        }
        else if (string.IsNullOrWhiteSpace(mainheads.Type))
        {
            throw new Exception("Cannot Proceed Without Type.");
        }
        else if (string.IsNullOrWhiteSpace(mainheads.MainGroupName))
        {
            throw new Exception("Cannot Proceed Without Main Group Name.");
        }

        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"
                    INSERT INTO GroupsForExpenseReport
                    (GroupName, SortNo, Type, MainGroupName, UserName, MachineName)
                    VALUES 
                    (@GroupName, @SortNo, @Type, @MainGroupName, @UserName, @MachineName)";
                var parameters = new
                {
                    GroupName = mainheads.GroupName,
                    SortNo = mainheads.SortNo,
                    Type = mainheads.Type,
                    MainGroupName = mainheads.MainGroupName,
                    UserName = mainheads.UserName,
                    MachineName = mainheads.MachineName,
                };
                int isavedrecords = await db.ExecuteAsync(sql, parameters);
                return isavedrecords > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Save: " + ex.Message);
            }
        }
    }

    public async Task<bool> EditGroupsAccounts(GroupsAccountsModel mainhead, long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql =
                    @"UPDATE GroupsForExpenseReport SET GroupName=@GroupName,Type=@Type,
                    MainGroupName=@MainGroupName
                    WHERE EntryID=@entryid";

                var parameters = new
                {
                    entryid = @entryid,
                    GroupName = mainhead.GroupName,
                    Type = mainhead.Type,
                    MainGroupName = mainhead.MainGroupName
                };
                int records = await db.ExecuteAsync(sql, parameters);
                return records > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Update: " + ex.Message);
            }
        }
    }

    public async Task<bool> DeleteGroupsAccounts(long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"DELETE FROM GroupsForExpenseReport WHERE EntryID=@entryid";
                int rowsAffected = await db.ExecuteAsync(sql, new { entryid });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Delete: " + ex.Message);
            }
        }
    }

    ////////////////////// ACCOUNTS HEAD MANUAL ENTRY ////////////////////////
    public async Task<List<AccountsHeadManualEntryModel>> GetAccountsHeadManualEntries()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var cond = "";
                cond = cond + " ORDER BY VAccountHeads_ManualEntry.DT ASC";
                string sql = @"
                    SELECT EntryID,DT,GroupName,Description,Notes,Amount,IsDetail,AccGroup_RefID FROM VAccountHeads_ManualEntry " + cond;

                    var entrieslist = await db.QueryAsync<AccountsHeadManualEntryModel>(sql);
                    return entrieslist.ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching : {ex.Message}");
            throw;
        }
    }
    public async Task<bool> DeleteAccountsHeadManualEntries(long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"DELETE FROM AccountHeads_ManualEntry WHERE EntryID=@entryid";
                int rowsAffected = await db.ExecuteAsync(sql, new { entryid });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Delete: " + ex.Message);
            }
        }
    }
    public async Task<bool> EditAccountsHeadManualEntries(AccountsHeadManualEntryModel mainhead, long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = "";

                sql =
                    @"UPDATE AccountHeads_ManualEntry SET 
                    Description=@Description,
                    Notes=@Notes,
                    Amount=@Amount,
                    DT=@DT
                    WHERE EntryID=@entryid";

                var parameters = new
                {
                    entryid = @entryid,
                    Description = mainhead.Description,
                    Notes = mainhead.Notes,
                    Amount = mainhead.Amount,
                    DT=mainhead.DT
                };
                int records = await db.ExecuteAsync(sql, parameters);
                return records > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Update: " + ex.Message);
            }
        }
    }

    public async Task<bool> SaveAccountsHeadManualEntries(AccountsHeadManualEntryModel mainheads)
    {
        if (mainheads.AccGroup_RefID == 0)
        {
            throw new Exception("Cannot Proceed Without Sub Group.");
        }

        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"
                    INSERT INTO AccountHeads_ManualEntry
                    (DT,AccGroup_RefID,Description,Notes,Amount,UserName,MachineName)
                    VALUES 
                    (@DT,@AccGroup_RefID,@Description,@Notes,@Amount,@UserName,@MachineName)";
                var parameters = new
                {
                    DT = mainheads.DT,
                    AccGroup_RefID = mainheads.AccGroup_RefID,
                    Description = mainheads.Description,
                    Notes = mainheads.Notes,
                    Amount = mainheads.Amount,
                    UserName = mainheads.UserName,
                    MachineName = mainheads.MachineName,
                };
                int isavedrecords = await db.ExecuteAsync(sql, parameters);
                return isavedrecords > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Save: " + ex.Message);
            }
        }
    }

    public async Task<List<SubHeadsListModel>> GetSubHeads(long MainGroup_refID)
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string cond= "";
                object? parameters = null;
                if (MainGroup_refID>0)
                { 
                    cond = " WHERE Type=@MainGroup_refID ";
                    parameters = new { MainGroup_refID };
                }
                cond += " ORDER BY GroupName ASC";

                string sql = @"SELECT EntryID,GroupName FROM GroupsForFinancialReport_Sub" + cond;

                return (await db.QueryAsync<SubHeadsListModel>(sql , parameters)).ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching: {ex.Message}");
            throw;
        }
    }

    ////////////////////// ACCOUNTS HEAD MANUAL ENTRY SHARE CAPITAL ////////////////////////
    public async Task<List<AccountsHeadManualEntry_SC_Model>> GetAccountsHeadManualEntries_SC()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var cond = "";
                cond = cond + " ORDER BY VAccountHeads_ShareCapital.DT ASC";
                string sql = @"
                    SELECT EntryID,DT,GroupName,Description,Notes,TotalAmount,AccGroup_RefID,ShareRate,ShareValue FROM VAccountHeads_ShareCapital " + cond;

                var entrieslist = await db.QueryAsync<AccountsHeadManualEntry_SC_Model>(sql);
                return entrieslist.ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching : {ex.Message}");
            throw;
        }
    }
    public async Task<bool> DeleteAccountsHeadManualEntries_SC(long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"DELETE FROM AccountHeads_ShareCapital WHERE EntryID=@entryid";
                int rowsAffected = await db.ExecuteAsync(sql, new { entryid });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Delete: " + ex.Message);
            }
        }
    }
    public async Task<bool> EditAccountsHeadManualEntries_SC(AccountsHeadManualEntry_SC_Model mainhead, long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = "";

                sql =
                    @"UPDATE AccountHeads_ShareCapital SET 
                    DT=@DT,Description=@Description,Notes=@Notes,ShareValue=@ShareValue,ShareRate=@ShareRate,TotalAmount=@TotalAmount
                    WHERE EntryID=@entryid";

                var parameters = new
                {
                    entryid = @entryid,
                    DT=mainhead.DT,
                    Description = mainhead.Description,
                    Notes = mainhead.Notes,
                    ShareValue=mainhead.ShareValue,
                    ShareRate=mainhead.ShareRate,
                    TotalAmount = mainhead.TotalAmount,
                };
                int records = await db.ExecuteAsync(sql, parameters);
                return records > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Update: " + ex.Message);
            }
        }
    }

    public async Task<bool> SaveAccountsHeadManualEntries_SC(AccountsHeadManualEntry_SC_Model mainheads)
    {

        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"
                    INSERT INTO AccountHeads_ShareCapital
                    (DT,AccGroup_RefID,Description,Notes,ShareValue,ShareRate,TotalAmount,UserName,MachineName)
                    VALUES 
                    (@DT,@AccGroup_RefID,@Description,@Notes,@ShareValue,@ShareRate,@TotalAmount,@UserName,@MachineName)";
                
                var parameters = new
                {
                    DT = mainheads.DT,
                    AccGroup_RefID = mainheads.AccGroup_RefID,
                    Description = mainheads.Description,
                    Notes = mainheads.Notes,
                    TotalAmount = mainheads.TotalAmount,
                    ShareRate = mainheads.ShareRate,
                    ShareValue = mainheads.ShareValue,
                    UserName = mainheads.UserName,
                    MachineName = mainheads.MachineName
                };
                int isavedrecords = await db.ExecuteAsync(sql, parameters);
                return isavedrecords > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Save: " + ex.Message);
            }
        }
    }

    ////////////////////// REMUNIRATION AND OTHER BENEFITS ////////////////////////
    public async Task<List<RemunerationandOtherBenefitsModel>> GetRemunerationandOtherBenefitsEntries()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var cond =" ORDER BY VAccountHeads_OtherBenefits.DT ASC";
                string sql = @"
                    SELECT EntryID,DT,GroupName,Description,Notes,CEAmount,AccGroup_RefID,DirAmount,ExeAmount FROM VAccountHeads_OtherBenefits " + cond ;

                var entrieslist = await db.QueryAsync<RemunerationandOtherBenefitsModel>(sql);
                return entrieslist.ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching : {ex.Message}");
            throw;
        }
    }
    public async Task<bool> DeleteRemunerationandOtherBenefitsEntries(long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"DELETE FROM AccountHeads_OtherBenefits WHERE EntryID=@entryid";
                int rowsAffected = await db.ExecuteAsync(sql, new { entryid });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Delete: " + ex.Message);
            }
        }
    }
    public async Task<bool> EditRemunerationandOtherBenefitsEntries(RemunerationandOtherBenefitsModel mainhead, long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = "";

                sql =
                    @"UPDATE AccountHeads_OtherBenefits SET 
                    DT=@DT,Description=@Description,Notes=@Notes,CEAmount=@CEAmount,DirAmount=@DirAmount,ExeAmount=@ExeAmount
                    WHERE EntryID=@entryid";

                var parameters = new
                {
                    entryid = @entryid,
                    DT = mainhead.DT,
                    Description = mainhead.Description,
                    Notes = mainhead.Notes,
                    CEAmount = mainhead.CEAmount,
                    DirAmount = mainhead.DirAmount,
                    ExeAmount = mainhead.ExeAmount,
                };
                int records = await db.ExecuteAsync(sql, parameters);
                return records > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Update: " + ex.Message);
            }
        }
    }

    public async Task<bool> SaveRemunerationandOtherBenefitsEntries(RemunerationandOtherBenefitsModel mainheads)
    {

        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"
                    INSERT INTO AccountHeads_OtherBenefits
                    (DT,AccGroup_RefID,Description,Notes,CEAmount,DirAmount,ExeAmount,UserName,MachineName)
                    VALUES 
                    (@DT,@AccGroup_RefID,@Description,@Notes,@CEAmount,@DirAmount,@ExeAmount,@UserName,@MachineName)";

                var parameters = new
                {
                    DT = mainheads.DT,
                    AccGroup_RefID = mainheads.AccGroup_RefID,
                    Description = mainheads.Description,
                    Notes = mainheads.Notes,
                    CEAmount = mainheads.CEAmount,
                    DirAmount = mainheads.DirAmount,
                    ExeAmount = mainheads.ExeAmount,
                    UserName = mainheads.UserName,
                    MachineName = mainheads.MachineName
                };
                int isavedrecords = await db.ExecuteAsync(sql, parameters);
                return isavedrecords > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Save: " + ex.Message);
            }
        }
    }


    ////////////////////// MANUAL ENTRY FOR NOTES ////////////////////////
    public async Task<List<ManualEntryForNotesModel>> GetManualEntryForNotesList()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var cond = " ORDER BY VAccountHeads_ManualEntry_ToShow.DT ASC";
                string sql = @"
                    SELECT EntryID,DT,GroupName,Description,Notes,AccGroup_RefID,Amount,Marked FROM VAccountHeads_ManualEntry_ToShow " + cond;

                var entrieslist = await db.QueryAsync<ManualEntryForNotesModel>(sql);
                return entrieslist.ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching : {ex.Message}");
            throw;
        }
    }
    public async Task<bool> DeleteManualEntryForNotes(long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"DELETE FROM AccountHeads_ManualEntry_ToShow WHERE EntryID=@entryid";
                int rowsAffected = await db.ExecuteAsync(sql, new { entryid });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Delete: " + ex.Message);
            }
        }
    }
    public async Task<bool> EditManualEntryForNotes(ManualEntryForNotesModel mainhead, long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = "";

                sql =
                    @"UPDATE AccountHeads_ManualEntry_ToShow SET 
                    DT=@DT,Description=@Description,Notes=@Notes,Amount=@Amount,Marked=@Marked
                    WHERE EntryID=@entryid";

                var parameters = new
                {
                    entryid = @entryid,
                    DT = mainhead.DT,
                    Description = mainhead.Description,
                    Notes = mainhead.Notes,
                    Amount = mainhead.Amount,
                    Marked = mainhead.Marked
                };
                int records = await db.ExecuteAsync(sql, parameters);
                return records > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Update: " + ex.Message);
            }
        }
    }

    public async Task<bool> SaveManualEntryForNotes(ManualEntryForNotesModel mainheads)
    {

        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"
                    INSERT INTO AccountHeads_ManualEntry_ToShow
                    (DT,AccGroup_RefID,Description,Notes,Amount,UserName,MachineName,Marked)
                    VALUES 
                    (@DT,@AccGroup_RefID,@Description,@Notes,@Amount,@UserName,@MachineName,@Marked)";

                var parameters = new
                {
                    DT = mainheads.DT,
                    AccGroup_RefID = mainheads.AccGroup_RefID,
                    Description = mainheads.Description,
                    Notes = mainheads.Notes,
                    Amount = mainheads.Amount,
                    UserName = mainheads.UserName,
                    MachineName = mainheads.MachineName,
                    Marked=mainheads.Marked
                };
                int isavedrecords = await db.ExecuteAsync(sql, parameters);
                return isavedrecords > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Save: " + ex.Message);
            }
        }
    }

    ////////////////////// SUB HEADS MANUAL ENTRY ////////////////////////
    public async Task<List<SubHeadsManualEntryModel>> GetSubHeadsManualEntriesList()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var cond = " ORDER BY T1.SrNo ASC";

                //cond = "WHERE T1.DTFrom BETWEEN @DTFrom AND @DTTo ";
                //cond = cond + " ORDER BY T1.SrNo ASC";
                string sql = @"
                    SELECT EntryID,AccGroup_RefID,GroupName,SubNotes,DTFrom,DTTo,SrNo,Description,Note AS Notes,
                           OpeningBalance,ClosingBalance,Marked,MainAccGroup_RefID,MainGroupName,MainNoteNo
                           FROM VSubHeadsManualEntry_Accounts T1 " + cond;

                var entrieslist = await db.QueryAsync<SubHeadsManualEntryModel>(sql);
                return entrieslist.ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching : {ex.Message}");
            throw;
        }
    }
    public async Task<bool> DeleteSubHeadsManualEntry(long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"DELETE FROM SubHeadsManualEntry_Accounts WHERE EntryID=@entryid";
                int rowsAffected = await db.ExecuteAsync(sql, new { entryid });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Delete: " + ex.Message);
            }
        }
    }
    public async Task<bool> EditSubHeadsManualEntries(SubHeadsManualEntryModel mainhead, long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = "";

                sql =
                    @"UPDATE SubHeadsManualEntry_Accounts SET 
                        SrNo=@SrNo,Description=@Description,Note=@Notes,OpeningBalance=@OpeningBalance,
                        ClosingBalance=@ClosingBalance,Marked=@Marked
                        WHERE EntryID=@entryid";

                var parameters = new
                {
                    entryid = @entryid,
                    SrNo = mainhead.SrNo,
                    Description = mainhead.Description,
                    Notes = mainhead.Notes,
                    OpeningBalance = mainhead.OpeningBalance,
                    ClosingBalance = mainhead.ClosingBalance,
                    Marked = mainhead.Marked
                };
                int records = await db.ExecuteAsync(sql, parameters);
                return records > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Update: " + ex.Message);
            }
        }
    }

    public async Task<bool> SaveSubHeadsManualEntries(SubHeadsManualEntryModel mainheads)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"
                    INSERT INTO SubHeadsManualEntry_Accounts
                    (AccGroup_RefID,DTFrom,DTTo,SrNo,Description,Note,OpeningBalance,ClosingBalance,Marked,UserName,MachineName,MainGroup_RefID)
                    VALUES 
                    (@AccGroup_RefID,@DTFrom,@DTTo,@SrNo,@Description,@Notes,@OpeningBalance,@ClosingBalance,@Marked,@UserName,@MachineName,@MainAccGroup_RefID)";

                var parameters = new
                {
                    AccGroup_RefID = mainheads.AccGroup_RefID,
                    DTFrom = mainheads.DTFrom,
                    DTTo = mainheads.DTTo,
                    SrNo = mainheads.SrNo,
                    Description = mainheads.Description,
                    Notes = mainheads.Notes,
                    OpeningBalance = mainheads.OpeningBalance,
                    ClosingBalance = mainheads.ClosingBalance,
                    Marked = mainheads.Marked,
                    MainAccGroup_RefID = mainheads.MainAccGroup_RefID,
                    UserName = mainheads.UserName,
                    MachineName = mainheads.MachineName,
                };
                int isavedrecords = await db.ExecuteAsync(sql, parameters);
                return isavedrecords > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Save: " + ex.Message);
            }
        }
    }



    ////////////////////// VOUCHERS ADJUSTMENT ////////////////////////
    public async Task<List<Vouchers_Adjustment_Transaction_Model>> GetTransactionData_Adjustment(DateTime DTFrom, DateTime DTTo)
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {

                var strcondition = "WHERE T1.DT BETWEEN @DTFrom AND @DTTo";
                strcondition = strcondition + " ORDER BY T1.DT ASC";

                string sql = @"
                               SELECT EntryID,DT,VchrNo,AccNo,AccTitle,Description,
                               Debit,Credit,RefID,UserName,MachineName,DTEntry
                               FROM VVouchers_Detail_Adj T1 " + strcondition;

                var listdata = await db.QueryAsync<Vouchers_Adjustment_Transaction_Model>
                    (sql, new { DTFrom ,DTTo, strcondition  });
                return listdata.ToList();

            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching data: {ex.Message}");
            throw;
        }
    }
    public async Task<bool> DeleteVoucher_Adjustment(string vchrno)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"DELETE FROM Vouchers_Main_Adj WHERE vchrno=@vchrno";
                int rowsAffected = await db.ExecuteAsync(sql, new { vchrno });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Delete: " + ex.Message);
            }
        }
    }
    public async Task<string> GetNextVchrNo_Adjustment(DateTime OnDate, string VoucherType)
    {
        string sql = "SELECT [dbo].[GetNextVchrNo_Adjustment](@DT,@VchrType)";
        var parameters = new
        {
            DT = OnDate,
            VchrType = VoucherType
        };
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            return await db.QueryFirstOrDefaultAsync<string>(sql, parameters);
        }
    }
    public async Task<string> SaveAdjustmentVoucher(Vouchers_Main_Model voucher)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            db.Open();
            using (var transaction = db.BeginTransaction())
            {
                try
                {
                    long nextentryid = await db.QuerySingleAsync<long>(
                        "dbo.Insert_Adjustment_Voucher_SP",
                        new
                        {
                            DT = voucher.DT,
                            VchrNo = voucher.VchrNo,
                            Remarks = voucher.Remarks,
                            UserName = voucher.UserName,
                            MachineName = voucher.MachineName
                        },
                        commandType: CommandType.StoredProcedure, transaction: transaction
                    );

                    long? returnSNo = null;
                    voucher.EntryID = nextentryid;
                    //returnSNo = await InsertVoucherAdjustmentDetailsAsync(db, transaction, voucher, nextentryid);
                    if (voucher.LineItems != null && voucher.LineItems.Any())
                    {
                        foreach (var detail in voucher.LineItems)
                        {
                            //detail.Balance = (detail.Debit - detail.Credit);
                            detail.AccNo = detail.AccNo;
                            await InsertVoucherAdjustmentDetailsAsync(db, transaction, detail, nextentryid, voucher.VchrNo);
                        }
                    }
                    transaction.Commit();
                    return nextentryid.ToString();
                }
                catch (Exception ex)
                {
                    {
                        transaction.Rollback();
                        _logger.LogError(ex,"Error: {ErrorMessage}",ex.Message);
                        throw;
                    }
                }
            }
        }
    }
    private async Task<long?> InsertVoucherAdjustmentDetailsAsync(IDbConnection db, IDbTransaction transaction, Vouchers_Detail_Model voucherdetail, long voucherentryid,string vchrno)
    {
        if (voucherdetail == null)
            return null;

        long lastentryid = await db.QuerySingleAsync<long>(
            "dbo.Insert_Adjustment_Voucher_Detail_SP",
            new
            {
                RefID = voucherentryid,
                VchrNo = vchrno,
                AccNo = voucherdetail.AccNo,
                Description = voucherdetail.Description,
                Debit = voucherdetail.Debit,
                Credit = voucherdetail.Credit,
                Balance = voucherdetail.Balance,
            },
            commandType: CommandType.StoredProcedure,
            transaction: transaction
        );
        return lastentryid;
    }

    ////////////////////// PROFIT AND LOSS HEADS ////////////////////////
    public async Task<List<MainHeadsListModel>> GetAllHeads()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string cond = "";
                cond = " ORDER BY VGroupsForFinancialReport_All.GroupName";

                string sql = @"SELECT EntryID,GroupName +' { ' + 'Note No. ' + Note + ' } ' GroupName FROM VGroupsForFinancialReport_All" + cond;
                return (await db.QueryAsync<MainHeadsListModel>(sql)).ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching: {ex.Message}");
            throw;
        }
    }
    public async Task<List<ProfitandLossHeadsDetailToLoadModel>> GetProfitandLossHeadsList()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var cond = " ORDER BY T1.SrNo ASC";
                string sql = @"
                    SELECT EntryID,SrNo,GroupName,Description,Note,IsTotal,Marked,NoteNo,IsMain,RefID FROM VProfitnLossHeads_Detail T1 " + cond;

                var entrieslist = await db.QueryAsync<ProfitandLossHeadsDetailToLoadModel>(sql);
                return entrieslist.ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching : {ex.Message}");
            throw;
        }
    }

    public async Task<string> SaveProfitandLossHeads(ProfitandLossHeads_Main_Model main)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            db.Open();
            using (var transaction = db.BeginTransaction())
            {
                try
                {
                    string mainSql = @"
                    INSERT INTO ProfitnLossHeads_Main 
                    (Description, SrNo, IsTotal, Marked, NoteNo) 
                    OUTPUT INSERTED.EntryID
                    VALUES 
                    (@Description, @SrNo, @IsTotal, @Marked, @NoteNo)";

                    long nextentryid = await db.QuerySingleAsync<long>(mainSql, new
                    {
                        Description = main.Description,
                        SrNo = main.SrNo,
                        IsTotal = main.IsTotal,
                        Marked = main.Marked,
                        NoteNo = main.NoteNo
                    }, transaction: transaction);

                    main.EntryID = nextentryid;

                    // 2. Insert Detail Records
                    if (main.DetailItems != null && main.DetailItems.Any())
                    {
                        var idList = string.Join(",", main.DetailItems.Select(x => x.Group_RefID));
                        string detailSql = $@"
                            INSERT INTO ProfitnLossHeads_Detail (RefID, Group_RefID, IsMain) 
                            SELECT @MasterID, EntryID, MainHead 
                            FROM VGroupsForFinancialReport_ALL 
                            WHERE EntryID IN ({idList})";

                        await db.ExecuteAsync(detailSql, new
                        {
                            MasterID = nextentryid
                        }, transaction: transaction);
                    }

                    transaction.Commit();
                    return nextentryid.ToString();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    _logger.LogError(ex, "Error On Save: {ErrorMessage}", ex.Message);
                    throw;
                }
            }
        }
    }

    public async Task<bool> DeleteProfitandLoss(long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"DELETE FROM ProfitnLossHeads_Main WHERE EntryID=@entryid";
                int rowsAffected = await db.ExecuteAsync(sql, new { entryid });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Delete: " + ex.Message);
            }
        }
    }

    public async Task<bool> EditProfitandLossHeads(ProfitandLossHeads_Main_Model mainhead, long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql =
                    @"UPDATE ProfitnLossHeads_Main SET 
                             Description=@Description,
                             SrNo=@SrNo,
                             Marked=@Marked,
                             IsTotal=@IsTotal,
                             NoteNo=@NoteNo
                             WHERE EntryID=@entryid";

                var parameters = new
                {
                    entryid = @entryid,
                    Marked = mainhead.Marked,
                    IsTotal = mainhead.IsTotal,
                    Description = mainhead.Description,
                    NoteNo = mainhead.NoteNo,
                    SrNo=mainhead.SrNo
                };
                int records = await db.ExecuteAsync(sql, parameters);
                return records > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Update: " + ex.Message);
            }
        }
    }

    ////////////// BALANCE SHEET 

    public async Task<List<BalanceSheetHeadsDetailToLoadModel>> GetBalanceSheetHeadsList()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var cond = " ORDER BY T1.SrNo ASC";
                string sql = @"
                    SELECT EntryID,SrNo,GroupName,Description,Note,IsTotal,Marked,NoteNo,IsMain,RefID,AssetType,HeadType FROM VBalanceSheetHeads_Detail T1 " + cond;

                var entrieslist = await db.QueryAsync<BalanceSheetHeadsDetailToLoadModel>(sql);
                return entrieslist.ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching : {ex.Message}");
            throw;
        }
    }

    public async Task<string> SaveBalanceSheetHeads(BalanceSheet_Main_Model main)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            db.Open();
            using (var transaction = db.BeginTransaction())
            {
                try
                {
                    string mainSql = @"
                    INSERT INTO BalanceSheetHeads_Main
                    (Description, SrNo, IsTotal, Marked, NoteNo, AssetType,HeadType) 
                    OUTPUT INSERTED.EntryID
                    VALUES 
                    (@Description, @SrNo, @IsTotal, @Marked, @NoteNo, @AssetType, @HeadType)";

                    long nextentryid = await db.QuerySingleAsync<long>(mainSql, new
                    {
                        Description = main.Description,
                        SrNo = main.SrNo,
                        IsTotal = main.IsTotal,
                        Marked = main.Marked,
                        NoteNo = main.NoteNo,
                        AssetType=main.AssetType,
                        HeadType = main.HeadType
                    }, transaction: transaction);

                    main.EntryID = nextentryid;

                    // 2. Insert Detail Records
                    if (main.DetailItems != null && main.DetailItems.Any())
                    {
                        var idList = string.Join(",", main.DetailItems.Select(x => x.Group_RefID));
                        string detailSql = $@"
                            INSERT INTO BalanceSheetHeads_Detail (RefID, Group_RefID, IsMain) 
                            SELECT @MasterID, EntryID, MainHead 
                            FROM VGroupsForFinancialReport_ALL 
                            WHERE EntryID IN ({idList})";

                        await db.ExecuteAsync(detailSql, new
                        {
                            MasterID = nextentryid
                        }, transaction: transaction);
                    }

                    transaction.Commit();
                    return nextentryid.ToString();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    _logger.LogError(ex, "Error On Save: {ErrorMessage}", ex.Message);
                    throw;
                }
            }
        }
    }

    public async Task<bool> DeleteBalanceSheetHeads(long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"DELETE FROM BalanceSheetHeads_Main WHERE EntryID=@entryid";
                int rowsAffected = await db.ExecuteAsync(sql, new { entryid });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Delete: " + ex.Message);
            }
        }
    }

    public async Task<bool> EditBalanceSheetHeads(BalanceSheet_Main_Model mainhead, long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql =
                    @"UPDATE BalanceSheetHeads_Main SET 
                             Description=@Description,
                             SrNo=@SrNo,
                             Marked=@Marked,
                             IsTotal=@IsTotal,
                             NoteNo=@NoteNo,
                             AssetType=@AssetType,
                             HeadType=@HeadType
                             WHERE EntryID=@entryid";

                var parameters = new
                {
                    entryid = @entryid,
                    Marked = mainhead.Marked,
                    IsTotal = mainhead.IsTotal,
                    Description = mainhead.Description,
                    NoteNo = mainhead.NoteNo,
                    SrNo = mainhead.SrNo,
                    AssetType = mainhead.AssetType,
                    HeadType = mainhead.HeadType
                };
                int records = await db.ExecuteAsync(sql, parameters);
                return records > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Update: " + ex.Message);
            }
        }
    }

    ////////////////// FINANCIAL STATEMENT

    public async Task<List<FinancialStatementHeadsModel>> GetFinancialStatementHeadsList()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var cond = " ORDER BY T1.MainHeadNote,T1.EntryType";
                string sql = @"
                    SELECT EntryID,RefID,GroupName,Note,MainHead,MainHeadNote AS MainNoteNo,ShowInBS,ShowInNTA FROM VFinancialStatementHeads T1 " + cond;

                var entrieslist = await db.QueryAsync<FinancialStatementHeadsModel>(sql);
                return entrieslist.ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching : {ex.Message}");
            throw;
        }
    }

    public async Task<string> SaveFinancialStatementHeads(string username, string machinename)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            db.Open();
            using (var transaction = db.BeginTransaction())
            {
                try
                {
                    string detailSql = @"
                    INSERT INTO FinancialStatementHeads (RefID, IsMain, UserName, MachineName)
                    SELECT EntryID, MainHead, @UserName, @MachineName 
                    FROM VGroupsForFinancialReport_All 
                    WHERE MainHead IN ( 0 , 1)
                    AND EntryID NOT IN (SELECT RefID FROM FinancialStatementHeads WHERE IsMain IN ( 0, 1) )";

                    int rowsAffected = await db.ExecuteAsync(detailSql, new
                    {
                        UserName = username,
                        MachineName = machinename
                    }, transaction: transaction);

                    transaction.Commit();

                    return rowsAffected.ToString();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    _logger.LogError(ex, "Error On Save: {ErrorMessage}", ex.Message);
                    throw;
                }
            }
        }
    }

    public async Task<bool> EditFinancialStatementHeads(FinancialStatementHeadsModel mainhead, int itype,long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = "";
                int records = 0;
                if (itype==1)
                { 
                sql =
                    @"UPDATE FinancialStatementHeads SET 
                            ShowinBS=@ShowinBS
                            WHERE EntryID=@entryid";
                    var parameters = new
                    {
                        entryid = @entryid,
                        ShowInBS = mainhead.ShowInBS
                    };
                    records = await db.ExecuteAsync(sql, parameters);
                }
                else
                {
                    sql =
                    @"UPDATE FinancialStatementHeads SET 
                        ShowInNTA=@ShowInNTA
                        WHERE EntryID=@entryid";

                    var parameters = new
                    {
                        entryid = @entryid,
                        ShowInNTA = mainhead.ShowInNTA
                    };
                    records = await db.ExecuteAsync(sql, parameters);
                }
                return records > 0;

            }
            catch (Exception ex)
            {
                throw new Exception("Error On Update: " + ex.Message);
            }
        }
    }

    ////////////////// CASH FLOW HEADS
    public async Task<List<CashFlowMainHeadsModel>> GetCashFlowMainHeadsList()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var cond = " ORDER BY T1.EntryID ASC";
                string sql = @"
                    SELECT EntryID,Description,Notes,Marked,CFHM_RefID,CFMainHead,CFH_RefID,CF_HeadType FROM VCashFlowHeadsForStatement T1 " + cond;

                var entrieslist = await db.QueryAsync<CashFlowMainHeadsModel>(sql);
                return entrieslist.ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching : {ex.Message}");
            throw;
        }
    }
    public async Task<bool> SaveCashFlowMainHeads(CashFlowMainHeadsModel cashflow)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            db.Open();
            using (var transaction = db.BeginTransaction())
            {
                try
                {
                    string detailSql = @"
                    INSERT INTO CashFlowHeadsForStatement (Description, CFHM_RefID, CFH_RefID, Notes)
                    VALUES (@Description, @CFHM_RefID, @CFH_RefID, @Notes)";

                    int rowsAffected = await db.ExecuteAsync(detailSql, new
                    {
                        Description = cashflow.Description,
                        CFHM_RefID = Convert.ToSingle(cashflow.CFHM_RefID),
                        CFH_RefID = Convert.ToSingle(cashflow.CFH_RefID),
                        Notes = cashflow.Notes
                    }, transaction: transaction);

                    transaction.Commit();
                    return rowsAffected > 0;
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    _logger?.LogError(ex, "Error On Save: {ErrorMessage}", ex.Message);

                    throw;
                }
            }
        }
    }
    public async Task<bool> EditCashFlowMainHeads(CashFlowMainHeadsModel mainhead, int itype, long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = "";
                int records = 0;
                if (itype == 1)
                {
                    sql =
                        @"UPDATE CashFlowHeadsForStatement SET 
                            Marked=@Marked,
                            Description=@Description,
                            Notes=@Notes
                            WHERE EntryID=@entryid";
                    var parameters = new
                    {
                        entryid = @entryid,
                        Description = mainhead.Description,
                        Notes = mainhead.Notes,
                        Marked=mainhead.Marked
                    };
                    records = await db.ExecuteAsync(sql, parameters);
                }
                else
                {
                    sql =
                    @"UPDATE CashFlowHeadsForStatement SET 
                        Marked=@Marked
                        WHERE EntryID=@entryid";

                    var parameters = new
                    {
                        entryid = @entryid,
                        Marked = mainhead.Marked
                    };
                    records = await db.ExecuteAsync(sql, parameters);
                }
                return records > 0;

            }
            catch (Exception ex)
            {
                throw new Exception("Error On Update: " + ex.Message);
            }
        }
    }
    public async Task<bool> DeletCashFlowMainHeads(long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"DELETE FROM CashFlowHeadsForStatement WHERE EntryID=@entryid";
                int rowsAffected = await db.ExecuteAsync(sql, new { entryid });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Delete: " + ex.Message);
            }
        }
    }

    ///////////// NOW PRINT REPORTS
    public async Task GenerateBalanceSheetReportAsync(DateTime dtFrom, DateTime dtTo)
    {

        DateTime openYear2 = new DateTime(dtFrom.Year, 7, 1);
        DateTime closeYear2 = dtTo;
        DateTime closeYear1 = dtTo.AddYears(-1);
        DateTime openYear1 = openYear2.AddYears(-1);

        using IDbConnection db = new SqlConnection(_connectionString);
        db.Open();

        using var transaction = db.BeginTransaction();

        try
        {
            // 1. Clear previous data
            await db.ExecuteAsync("DELETE FROM BalanceSheet", transaction: transaction);

            // 2. Fetch Primary Data from SP
            var parameters1 = new
            {
                PrvDTStart = openYear2,
                PrvDTEnd = closeYear2,
                CurDTStart = openYear2,
                CurDTEnd = closeYear2
            };

            var primaryData = await db.QueryAsync<dynamic>(
                "BalanceSheet_SIAL_SP",
                parameters1,
                commandType: CommandType.StoredProcedure,
                transaction: transaction
            );

            // 3. Insert Primary Data
            foreach (var item in primaryData)
            {
                await InsertEntry(db, transaction, item, item.AccNo);
            }

            // 4. Fetch Detail Data from SP
            var parameters2 = new
            {
                PrvDTStart = openYear1,
                PrvDTEnd = closeYear1,
                CurDTStart = openYear2,
                CurDTEnd = closeYear2
            };

            var detailData = await db.QueryAsync<dynamic>(
                "BalanceSheet_Heads_Detail_SIAL_SP",
                parameters2,
                commandType: CommandType.StoredProcedure,
                transaction: transaction
            );

            // 5. Insert Detail Data (Force AccNo to "FROM HEAD")
            foreach (var item in detailData)
            {
                // Applying your VBA logic: forced 0s for specific amounts in detail
                await InsertEntry(db, transaction, item, "FROM HEAD", true);
            }

            transaction.Commit();
        }
        catch (Exception)
        {
            transaction.Rollback();
            throw;
        }
    }

    private async Task InsertEntry(IDbConnection db, IDbTransaction trans, dynamic item, string accNo, bool isDetailPart = false)
    {
        string sql = @"INSERT INTO BalanceSheet (AccNo, AccTitle, SubAccOf, OpeningBalance, PrvAmount, CurAmount, LastAmount, EntryId, GroupName, Notes, SubGroupName, SubEntryID, IsDetail) 
                       VALUES (@AccNo, @AccTitle, @SubAccOf, @OpeningBalance, @PrvAmount, @CurAmount, @LastAmount, @EntryId, @GroupName, @Notes, @SubGroupName, @SubEntryID, 0)";

        await db.ExecuteAsync(sql, new
        {
            AccNo = accNo,
            AccTitle = item.AccTitle ?? "",
            SubAccOf = item.SubAccOf ?? "",

            OpeningBalance = (decimal)(item.OpeningBalance ?? 0),
            
            PrvAmount = isDetailPart ? 0 : (decimal)(item.PrvAmount ?? 0),
            CurAmount = isDetailPart ? 0 : (decimal)(item.CurAmount ?? 0),
            LastAmount = (decimal)(item.LastBalance ?? 0),
            EntryId = item.EntryID?.ToString() ?? "",
            GroupName = item.GroupName ?? "",
            Notes = item.SortNo?.ToString() ?? "",
            SubGroupName = item.Subgroup ?? "",
            SubEntryID = item.SubEntryID?.ToString() ?? ""
        }, transaction: trans);
    }

    public async Task GenerateBalanceSheetReportAsync_Prv(DateTime dtFrom, DateTime dtTo)
    {

        DateTime openYear2 = new DateTime(dtFrom.Year, 7, 1);
        DateTime closeYear2 = dtTo;
        DateTime closeYear1 = dtTo.AddYears(-1);
        DateTime openYear1 = openYear2.AddYears(-1);

        using IDbConnection db = new SqlConnection(_connectionString);
        db.Open();

        using var transaction = db.BeginTransaction();

        try
        {
            // 1. Clear previous data
            await db.ExecuteAsync("DELETE FROM BalanceSheet_Prv", transaction: transaction);

            // 2. Fetch Primary Data from SP
            var parameters1 = new
            {
                PrvDTStart = openYear1,
                PrvDTEnd = closeYear1,
                CurDTStart = openYear1,
                CurDTEnd = closeYear1
            };

            var primaryData = await db.QueryAsync<dynamic>(
                "BalanceSheet_SIAL_SP",
                parameters1,
                commandType: CommandType.StoredProcedure,
                transaction: transaction
            );

            // 3. Insert Primary Data
            foreach (var item in primaryData)
            {
                await InsertEntry_Prv(db, transaction, item, item.AccNo);
            }

            // 4. Fetch Detail Data from SP
            var parameters2 = new
            {
                PrvDTStart = openYear1,
                PrvDTEnd = closeYear1,
                CurDTStart = openYear1,
                CurDTEnd = closeYear1
            };

            var detailData = await db.QueryAsync<dynamic>(
                "BalanceSheet_Heads_Detail_SIAL_SP",
                parameters2,
                commandType: CommandType.StoredProcedure,
                transaction: transaction
            );

            foreach (var item in detailData)
            {
                await InsertEntry_Prv(db, transaction, item, "FROM HEAD", true);
            }

            transaction.Commit();
        }
        catch (Exception)
        {
            transaction.Rollback();
            throw;
        }
    }

    private async Task InsertEntry_Prv(IDbConnection db, IDbTransaction trans, dynamic item, string accNo, bool isDetailPart = false)
    {
        string sql = @"INSERT INTO BalanceSheet_Prv(AccNo, AccTitle, SubAccOf, OpeningBalance, PrvAmount, CurAmount, LastAmount, EntryId, GroupName, Notes, SubGroupName, SubEntryID, IsDetail) 
                       VALUES (@AccNo, @AccTitle, @SubAccOf, @OpeningBalance, @PrvAmount, @CurAmount, @LastAmount, @EntryId, @GroupName, @Notes, @SubGroupName, @SubEntryID, 0)";

        await db.ExecuteAsync(sql, new
        {
            AccNo = accNo,
            AccTitle = item.AccTitle ?? "",
            SubAccOf = item.SubAccOf ?? "",
            OpeningBalance = (decimal)(item.OpeningBalance ?? 0),

            PrvAmount = isDetailPart ? 0 : (decimal)(item.PrvAmount ?? 0),
            CurAmount = isDetailPart ? 0 : (decimal)(item.CurAmount ?? 0),
            LastAmount = (decimal)(item.LastBalance ?? 0),
            EntryId = item.EntryID?.ToString() ?? "",
            GroupName = item.GroupName ?? "",
            Notes = item.SortNo?.ToString() ?? "",
            SubGroupName = item.Subgroup ?? "",
            SubEntryID = item.SubEntryID?.ToString() ?? ""
        }, transaction: trans);
    }

    public async Task<int> SaveNotesAttachments(PrintAttachmentsModel mainheads)
    {
        if (string.IsNullOrWhiteSpace(mainheads.Attachment_Name))
        {
            throw new Exception("Cannot Proceed Without Attachment Name.");
        }
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"
                    INSERT INTO AccountsNotes_Attachments
                    (Attachment_Name, Attachment_Type)
                    VALUES 
                    (@Attachment_Name, @Attachment_Type);
                    
                    SELECT CAST(SCOPE_IDENTITY() as int);
                    
                ";
                var parameters = new
                {
                    Attachment_Name = mainheads.Attachment_Name,
                    Attachment_Type = mainheads.Attachment_Type,
                };
                int isavedrecords = await db.QuerySingleAsync<int>(sql, parameters);
                return isavedrecords;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Save: " + ex.Message);
            }
        }
    }

    public async Task<List<PrintAttachmentsModel>> GetNotesAttachmentList()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var cond = " ORDER BY T1.Attachment_Name";
                string sql = @"
                    SELECT EntryID,Attachment_Name,Attachment_Type FROM AccountsNotes_Attachments T1 " + cond;

                var entrieslist = await db.QueryAsync<PrintAttachmentsModel>(sql);
                return entrieslist.ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching : {ex.Message}");
            throw;
        }
    }

    public async Task<List<CashFlowType_DD_Model>> GetCashFlowType_DD()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"SELECT EntryID,CF_HeadType FROM CashFlowHeads T1";
                sql += " ORDER BY T1.CF_HeadType ASC";
                return (await db.QueryAsync<CashFlowType_DD_Model>(sql)).ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching: {ex.Message}");
            throw;
        }
    }
    public async Task<List<CashFlowMain_DD_Model>> GetCashFlowMainHeads_DD()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"SELECT EntryID,CFMainHead FROM CashFlowMainHeads T1";
                sql += " ORDER BY T1.CFMainHead";
                return (await db.QueryAsync<CashFlowMain_DD_Model>(sql)).ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching: {ex.Message}");
            throw;
        }
    }

///////////// CASH FLOW REPORTS HEDS

    public async Task<List<CashFlowReportHeadsDetailToLoadModel>> GetCashFlowReportHeadsList()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var cond = " ORDER BY T1.SrNo ASC";
                string sql = @"
                    SELECT EntryID,RefID,Description,SrNo,Group_RefID,IsMain,GroupName,Notes,IsTotal,Marked FROM VCashFlowReportHeads_Detail T1 " + cond;

                var entrieslist = await db.QueryAsync<CashFlowReportHeadsDetailToLoadModel>(sql);
                return entrieslist.ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching : {ex.Message}");
            throw;
        }
    }

    public async Task<List<CashFlowHeadsListModel>> GetCashFlowHeadsList_DD()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string cond = "";
                cond = " ORDER BY VCashFlowHeadsForStatement.Description";

                string sql = @"SELECT EntryID,Description GroupName FROM VCashFlowHeadsForStatement" + cond;
                return (await db.QueryAsync<CashFlowHeadsListModel>(sql)).ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on fetching: {ex.Message}");
            throw;
        }
    }

    public async Task<string> SaveCashFlowReportHeads(CashFlowReportHeads_Main_Model main)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            db.Open();
            using (var transaction = db.BeginTransaction())
            {
                try
                {
                    string mainSql = @"
                    INSERT INTO CashFlowHeads_Main 
                    (Description,SrNo) 
                    OUTPUT INSERTED.EntryID
                    VALUES 
                    (@Description, @SrNo)";

                    long nextentryid = await db.QuerySingleAsync<long>(mainSql, new
                    {
                        Description = main.Description,
                        SrNo = main.SrNo
                    }, transaction: transaction);

                    main.EntryID = nextentryid;

                    // 2. Insert Detail Records
                    if (main.DetailItems != null && main.DetailItems.Any())
                    {
                        var idList = string.Join(",", main.DetailItems.Select(x => x.Group_RefID));
                        string detailSql = $@"
                            INSERT INTO CashFlowHeads_Detail (RefID,Group_RefID,IsMain) 
                            SELECT @MasterID, EntryID,0
                            FROM VCashFlowHeadsForStatement 
                            WHERE EntryID IN ({idList})";

                        await db.ExecuteAsync(detailSql, new
                        {
                            MasterID = nextentryid
                        }, transaction: transaction);
                    }

                    transaction.Commit();
                    return nextentryid.ToString();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    _logger.LogError(ex, "Error On Save: {ErrorMessage}", ex.Message);
                    throw;
                }
            }
        }
    }

    public async Task<bool> DeleteCashFlowReportHeads(long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql = @"DELETE FROM CashFlowHeads_Main WHERE EntryID=@entryid";
                int rowsAffected = await db.ExecuteAsync(sql, new { entryid });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Delete: " + ex.Message);
            }
        }
    }

    public async Task<bool> EditCashFlowReportHeads(CashFlowReportHeads_Main_Model mainhead, long entryid)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                string sql =
                    @"UPDATE ProfitnLossHeads_Main SET 
                             Description=@Description,
                             SrNo=@SrNo,
                             Marked=@Marked,
                             IsTotal=@IsTotal,
                             NoteNo=@NoteNo
                             WHERE EntryID=@entryid";

                var parameters = new
                {
                    entryid = @entryid,
                    Marked = mainhead.Marked,
                    IsTotal = mainhead.IsTotal,
                    Description = mainhead.Description,
                    NoteNo = mainhead.NoteNo,
                    SrNo = mainhead.SrNo
                };
                int records = await db.ExecuteAsync(sql, parameters);
                return records > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error On Update: " + ex.Message);
            }
        }
    }

    /////////// CAPITAL WORK IN PROGRESS
    public async Task ProcessCapitalWorkInProgress()
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                
            /////// CURRENT YEAR
                await db.ExecuteAsync("DELETE FROM CapitalWorkInProgress");

                string checkSql = "SELECT COUNT(*) FROM BalanceSheet WHERE SubAccOf IN ('21-002-001', '21-002-004')";
                int count = await db.ExecuteScalarAsync<int>(checkSql);

                if (count > 0)
                {
                    string moveSql = @"
                    INSERT INTO CapitalWorkInProgress SELECT * FROM BalanceSheet WHERE SubAccOf IN ('21-002-001', '21-002-004');
                    DELETE FROM BalanceSheet WHERE SubAccOf IN ('21-002-001', '21-002-004');";

                    await db.ExecuteAsync(moveSql);
                }

            /////// PRV YEAR
                await db.ExecuteAsync("DELETE FROM CapitalWorkInProgress_Prv");

                string checkPrvSql = "SELECT COUNT(*) FROM BalanceSheet_Prv WHERE SubAccOf IN ('21-002-001', '21-002-004')";
                int countPrv = await db.ExecuteScalarAsync<int>(checkPrvSql);

                if (countPrv > 0)
                {
                    string movePrvSql = @"
                    INSERT INTO CapitalWorkInProgress_Prv SELECT * FROM BalanceSheet_Prv WHERE SubAccOf IN ('21-002-001', '21-002-004');
                    DELETE FROM BalanceSheet_Prv WHERE SubAccOf IN ('21-002-001', '21-002-004');";

                    await db.ExecuteAsync(movePrvSql);
                }

            }
            catch (Exception ex)
            {
                throw new Exception("Error Processing WIP Data: " + ex.Message);
            }
        }
    }

    /////////// SAVE PROPERTY
    public async Task SaveProperty()
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            try
            {
                // 1. Current Year Calculation (BalanceSheet_Sub_1 & CapitalWorkInProgress)
                double dNetValue = await _db.getSingleDoubleValueasync("SUM(NetValue)", "BalanceSheet_Sub_1", "");
                double netValue_Acc = await _db.getSingleDoubleValueasync("SUM(LastAmount)", "CapitalWorkInProgress", "");

                double finalCurrentValue = dNetValue - netValue_Acc;

                // 2. Insert into BalanceSheet
                string sqlInsert = @"
                INSERT INTO BalanceSheet (AccNo, AccTitle, SubAccOf, LastAmount, EntryId, GroupName, Notes, SubGroupName, SubEntryId) 
                VALUES ('41-039-19001', 'Property plant and equipment', '41-039', @LastAmount, '1', 
                        'Property plant And equipment', 4, 'Property plant And equipment', '5000')";

                await db.ExecuteAsync(sqlInsert, new { LastAmount = finalCurrentValue });

                // 3. Previous Year Calculation (BalanceSheet_Sub_2 & CapitalWorkInProgress_Prv)
                double dNetValuePrv = await _db.getSingleDoubleValueasync("SUM(NetValue)", "BalanceSheet_Sub_2", "");
                double netValue_AccPrv = await _db.getSingleDoubleValueasync("SUM(LastAmount)", "CapitalWorkInProgress_Prv", "");

                double finalPrevValue = dNetValuePrv - netValue_AccPrv;

                // 4. Update BalanceSheet OpeningBalance
                string sqlUpdate = "UPDATE BalanceSheet SET OpeningBalance = @OpeningBalance WHERE AccNo = '41-039-19001'";

                await db.ExecuteAsync(sqlUpdate, new { OpeningBalance = finalPrevValue });
            }
            catch (Exception ex)
            {
                throw new Exception("Error Saving Property Data: " + ex.Message);
            }
        }
    }
    /////////// SAVE BALANCE SHEET TO PRINT
    public async Task SaveBalanceSheetToPrint()
    {
        try
        {
            // 1. Clear existing print table
            await _db.ExecuteAsync("DELETE FROM Print_BalanceSheet");

            // 2. SAVE PROPERTY TOTALS
            double dNetValuePrv = await _db.getSingleDoubleValueasync("SUM(NetValue)", "BalanceSheet_Sub_2", "");
            double dCapTotalPrv = await _db.getSingleDoubleValueasync("SUM(LastAmount)", "CapitalWorkInProgress_Prv", "");

            double dNetValueCur = await _db.getSingleDoubleValueasync("SUM(NetValue)", "BalanceSheet_Sub_1", "");
            double dCapTotalCur = await _db.getSingleDoubleValueasync("SUM(LastAmount)", "CapitalWorkInProgress", "");

            await InsertPrintRow("4", "Property Plant and Equipment", (dNetValuePrv - dCapTotalPrv), (dNetValueCur - dCapTotalCur), 1);

            // 3. Process standard rows by EntryID
            // Pattern: (Note, Title, EntryID, SortOrder)
            await ProcessStandardEntry("", "Long Term Security Deposits", 3, 2);
            await ProcessStandardEntry("", "Store And Spares", 4, 3);
            await ProcessStandardEntry("5", "Trade Debts", 5, 4);
            await ProcessStandardEntry("6", "Advances Prepayments and Other Receivables", 6, 5);
            await ProcessStandardEntry("7", "Cash and Bank Balances", 8, 6);
            await ProcessStandardEntry("8", "Share Capital", 9, 7);
            await ProcessStandardEntry("", "Share Premium", 10, 8);
            await ProcessStandardEntry("9", "General Reserve", 11, 9);
            await ProcessStandardEntry("", "UnAppropriated Profit", 12, 10);
            await ProcessStandardEntry("10", "Long Term Loans", 13, 11);
            await ProcessStandardEntry("11", "Liabilities Against Assets Subject To Finance Lease", 14, 12);
            await ProcessStandardEntry("12", "Deferred Liabilities", 15, 13);
            await ProcessStandardEntry("13", "Deferred Income", 16, 14);
            await ProcessStandardEntry("14", "Trade and Other Payables", 17, 15);
            await ProcessStandardEntry("", "Markup Accrued On Long Term Loans", 25, 16);
            await ProcessStandardEntry("15", "Current Portion Of Non-Current Liabilities", 26, 17);
            await ProcessStandardEntry("", "Provision For Taxation", 27, 18);
        }
        catch (Exception ex)
        {
            throw new Exception("Error in SaveBalanceSheetToPrint: " + ex.Message);
        }
    }

    // Helper to handle the repetitive EntryID logic
    private async Task ProcessStandardEntry(string note, string title, int entryId, int sortNo)
    {
        double prv = await _db.getSingleDoubleValueasync("SUM(LastAmount)", "BalanceSheet_Prv", $"WHERE EntryId={entryId}");
        double cur = await _db.getSingleDoubleValueasync("SUM(LastAmount)", "BalanceSheet", $"WHERE EntryId={entryId}");
        await InsertPrintRow(note, title, prv, cur, sortNo);
    }

    // Helper to execute the actual INSERT
    private async Task InsertPrintRow(string notes, string name, double opening, double closing, int sortNo)
    {
        string sql = @"INSERT INTO Print_BalanceSheet (Notes, MainGroupName, OpeningBalance, ClosingBalance, SortNo) 
                   VALUES (@Notes, @Name, @Opening, @Closing, @SortNo)";

        var parameters = new
        {
            Notes = notes,
            Name = name,
            Opening = opening,
            Closing = closing,
            SortNo = sortNo
        };

        await _db.ExecuteAsync(sql, parameters);
    }

    public async Task PrintIncomeAndExpenseReport()
    {
        try
        {
            // 1. Clear existing print table
            await _db.ExecuteAsync("DELETE FROM PrintIncomeANDExpense");

            // 2. Fetch main data
            string fetchSql = "SELECT * FROM IncomeANDExpense";
            var items = await _db.getListasync<dynamic>(fetchSql);

            foreach (var item in items)
            {
                // Calculation: Balance = Opening + Prv - Cur
                double opening = Convert.ToDouble(item.OpeningBalance ?? 0);
                double prv = Convert.ToDouble(item.PrvAmount ?? 0);
                double cur = Convert.ToDouble(item.CurAmount ?? 0);
                double balanceAmount = opening + prv - cur;

                // INSERT Row
                string insertSql = @"INSERT INTO PrintIncomeANDExpense (AccNo, AccTitle, SubAccOf, CurAmount, PrvAmount, EntryId, GroupName, Notes, SubGroupName) 
                                 VALUES (@AccNo, @AccTitle, @SubAccOf, @CurAmount, 0, @EntryId, @GroupName, @Notes, @SubGroupName)";

                await _db.ExecuteAsync(insertSql, new
                {
                    AccNo = item.AccNo,
                    AccTitle = item.AccTitle ?? "",
                    SubAccOf = item.SubAccOf ?? "",
                    CurAmount = balanceAmount,
                    EntryId = item.EntryID?.ToString() ?? "",
                    GroupName = item.GroupName ?? "",
                    Notes = item.Notes ?? "",
                    SubGroupName = item.SubGroupName ?? ""
                });

                // UPDATE Previous Amount from IncomeANDExpense_Prv
                double prvBalance = await _db.getSingleDoubleValueasync(
                    "OpeningBalance + PrvAmount - CurAmount",
                    "IncomeANDExpense_Prv",
                    $"WHERE AccNo='{item.AccNo}'"
                );

                await _db.ExecuteAsync("UPDATE PrintIncomeANDExpense SET PrvAmount=@val WHERE AccNo=@acc",
                                        new { val = prvBalance, acc = item.AccNo });
            }

            // 3. Depreciation Rows
            double depCur = await _db.getSingleDoubleValueasync("SUM(PrvAmount_Acc)", "BalanceSheet_Sub_1", "");
            double depPrv = await _db.getSingleDoubleValueasync("SUM(PrvAmount_Acc)", "BalanceSheet_Sub_2", "");

            await InsertPrintIncomeRow("01-01-01-01", "Depreciation", "01-01-01", depCur, depPrv, "111", "Depreciation", "4", "Depreciation");

            // 4. Tax Calculation (30%)
            double baseAmount = await _db.getSingleDoubleValueasync("SUM(CurAmount)", "PrintIncomeANDExpense", "WHERE Notes IN ('18','19','4','20','21')");
            double baseAmountPrv = await _db.getSingleDoubleValueasync("SUM(PrvAmount)", "PrintIncomeANDExpense", "WHERE Notes IN ('18','19','4','20','21')");

            double tax30 = Math.Round(baseAmount * 0.30, 0);
            double tax30Prv = Math.Round(baseAmountPrv * 0.30, 0);

            await InsertPrintIncomeRow("01-01-01-01", "Provision For Taxation", "01-01-01", tax30, tax30Prv, "112", "Provision For Taxation", "22", "Provision For Taxation");

            // 5. Tax Calculation (2.32%)
            double tax232 = Math.Round(baseAmount * 0.0232, 0);
            double tax232Prv = Math.Round(baseAmountPrv * 0.0232, 0);

            await InsertPrintIncomeRow("01-01-01-01", "Provision For Taxation", "01-01-01", tax232, tax232Prv, "113", "Provision For Taxation", "22", "Provision For Taxation");
        }
        catch (Exception ex)
        {
            throw new Exception("Error in PrintIncomeAndExpenseReport: " + ex.Message);
        }
    }

    // Helper to keep code clean
    private async Task InsertPrintIncomeRow(string accNo, string title, string subOf, double cur, double prv, string entry, string group, string notes, string subGroup)
    {
        string sql = @"INSERT INTO PrintIncomeANDExpense (AccNo, AccTitle, SubAccOf, CurAmount, PrvAmount, EntryId, GroupName, Notes, SubGroupName) 
                   VALUES (@AccNo, @AccTitle, @SubOf, @Cur, @Prv, @Entry, @Group, @Notes, @SubGroup)";

        await _db.ExecuteAsync(sql, new { AccNo = accNo, AccTitle = title, SubOf = subOf, Cur = cur, Prv = prv, Entry = entry, Group = group, Notes = notes, SubGroup = subGroup });
    }


    /////////////////// CASH FLOW STATEMENT REPORT
    public async Task SaveCashFlowStatementToPrint(DateTime dtFromPrv, DateTime dtToPrv, DateTime dtFromCur, DateTime dtToCur, bool bPrvYear = false)
    {
        try
        {
            // 1. Execute Stored Procedure using Dapper
            // Assuming your DBHelper has a method to get a list from a stored procedure
            var parameters = new
            {
                PrvDTStart = dtFromPrv,
                PrvDTEnd = dtToPrv,
                CurDTStart = dtFromCur,
                CurDTEnd = dtToCur
            };

            // Fetching the results of the SP into a list of dynamic objects
            var rsAccounts = await _db.getListasync<dynamic>("CashFlow_SIAL_SP", parameters);

            // 2. Clear table if not processing previous year
            if (!bPrvYear)
            {
                await _db.ExecuteAsync("DELETE FROM CashFlowStatement");
            }

            foreach (var item in rsAccounts)
            {
                // --- CALCULATION LOGIC ---
                // Round(Abs(Cur_Dr) - Abs(Prv_Dr))
                double dDiff_Dr = Math.Round(Math.Abs(Convert.ToDouble(item.ClBal_Cur_Dr ?? 0)) - Math.Abs(Convert.ToDouble(item.ClBal_Prv_Dr ?? 0)), 0);
                double dDiff_Cr = Math.Round(Math.Abs(Convert.ToDouble(item.ClBal_Cur_Cr ?? 0)) - Math.Abs(Convert.ToDouble(item.ClBal_Prv_Cr ?? 0)), 0);

                double dPL = 0;
                double dBS = 0;
                string type = item.Type?.ToString() ?? "";
                int cfRefId = Convert.ToInt32(item.CFHFS_RefID ?? 0);
                string accNo = item.AccNo?.ToString() ?? "";

                if (accNo == "21-001-036-009")
                {
                    if (System.Diagnostics.Debugger.IsAttached)
                    {
                        //System.Diagnostics.Debugger.Break();
                    }
                }

                // P&L Logic (Types 31, 33, 41-46)
                if (new[] { "31", "33", "41", "42", "43", "44", "45", "46" }.Contains(type))
                {
                    if (type == "31" || type == "33")
                    {
                        dPL = dDiff_Dr > 0 ? (dDiff_Cr - dDiff_Dr) : (dDiff_Cr - (-dDiff_Dr));
                    }
                    else // Types 41-46
                    {
                        dPL = (dDiff_Cr - dDiff_Dr) > 1 ? (dDiff_Cr - dDiff_Dr) : ((-dDiff_Dr) - dDiff_Cr);
                    }
                }

                // Balance Sheet Logic (Types 11, 12, 13, 21, 22)
                if (new[] { "11", "12", "13", "21", "22" }.Contains(type))
                {
                    dBS = dDiff_Cr - dDiff_Dr;
                }

                // --- SPECIAL ACCOUNT LOGIC (13-006, 12-007) ---
                if (cfRefId == 31 && accNo == "13-006")
                {
                    double dClosingBalance = await _db.getSingleDoubleValueasync("SUM(ISNULL(TotAmt,0))", "dbo.F_LeaseRent_DebitBalance", $"('{dtFromCur:yyyy-MM-dd}', '{dtToCur:yyyy-MM-dd}')");
                    await InsertCashFlowRow("Manual Added", item.AccTitle, 0, 0, 0, 0, 0, dClosingBalance, dPL, dClosingBalance, cfRefId, item.Marked_CF_Nve, item.Marked_CF, bPrvYear);
                }
                else if (cfRefId == 22) // Additions to property and equipment
                {
                    // Check if already added manually
                    int iCount = await _db.getSingleIntValueasync("COUNT(*)", "CashFlowStatement", $"WHERE AccNo='Manual Added' AND bPrvYear='{bPrvYear}' AND CFHFS_RefID={cfRefId}");

                    if (iCount == 0 && (accNo == "12-007" || accNo == "13-006"))
                    {
                        double dClosingBalance = await _db.getSingleDoubleValueasync("SUM(ISNULL(TotAmt,0))", "dbo.F_LeaseRent_CreditBalance", $"('{dtFromCur:yyyy-MM-dd}', '{dtToCur:yyyy-MM-dd}')");

                        await InsertCashFlowRow("Manual Added", "", 0, 0, 0, 0, 0, 0, 0, dClosingBalance, cfRefId, 1, 1, bPrvYear);

                        if (accNo == "12-007")
                        {
                            double dClosingBalance_Dr = -await _db.getSingleDoubleValueasync("SUM(ISNULL(TotAmt,0))", "dbo.F_LeaseRent_DebitBalance", $"('{dtFromCur:yyyy-MM-dd}', '{dtToCur:yyyy-MM-dd}')");
                            await InsertCashFlowRow("Manual Added", "", 0, 0, 0, 0, 0, 0, 0, dClosingBalance_Dr, cfRefId, 1, 1, bPrvYear);
                        }
                    }
                }

                // --- STANDARD ROW INSERT ---
                if (accNo != "12-007" && accNo != "13-006")
                {
                    await InsertCashFlowRow(
                        accNo,
                        item.AccTitle,
                        Convert.ToDouble(item.ClBal_Prv_Dr ?? 0),
                        Convert.ToDouble(item.ClBal_Prv_Cr ?? 0),
                        Convert.ToDouble(item.ClBal_Cur_Dr ?? 0),
                        Convert.ToDouble(item.ClBal_Cur_Cr ?? 0),
                        dDiff_Dr,
                        dDiff_Cr,
                        dPL,
                        dBS,
                        cfRefId,
                        item.Marked_CF_Nve,
                        item.Marked_CF,
                        bPrvYear
                    );
                }
            }
        }
        catch (Exception ex)
        {
            throw new Exception("CashFlow Error: " + ex.Message);
        }
    }

    // Helper method to handle the INSERT SQL
    private async Task InsertCashFlowRow(string accNo, string title, double prvDr, double prvCr, double curDr, double curCr, double diffDr, double diffCr, double pl, double bs, int refId, object marked, object markedCf, bool bPrv)
    {
        string sql = @"INSERT INTO CashFlowStatement (AccNo, AccTitle, ClosingBalance_Prv_Dr, ClosingBalance_Prv_Cr, ClosingBalance_Cur_Dr, ClosingBalance_Cur_Cr, Diff_Dr, Diff_Cr, PL, BS, CFHFS_RefID, Marked, Marked_CF, bPrvYear) 
                   VALUES (@AccNo, @Title, @PrvDr, @PrvCr, @CurDr, @CurCr, @DiffDr, @DiffCr, @PL, @BS, @RefId, @Marked, @MarkedCf, @BPrv)";

        await _db.ExecuteAsync(sql, new
        {
            AccNo = accNo,
            Title = title ?? "",
            PrvDr = prvDr,
            PrvCr = prvCr,
            CurDr = curDr,
            CurCr = curCr,
            DiffDr = diffDr,
            DiffCr = diffCr,
            PL = pl,
            BS = bs,
            RefId = refId,
            Marked = marked?.ToString() ?? "",
            MarkedCf = markedCf,
            BPrv = bPrv
        });
    }


}
