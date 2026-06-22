using Dapper;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Configuration;
using OfficeOpenXml.Sorting;
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
using System.Net.NetworkInformation;
using System.Reflection.PortableExecutable;
using System.Threading.Tasks;
using static System.Net.WebRequestMethods;

public class AccountsReportingDataAccess : IAccountReportingAccess
{
    private readonly string _connectionString;
    public AccountsReportingDataAccess(IConfiguration config)
    {
        _connectionString = config.GetConnectionString("DefaultConnection");
    }

    public async Task<List<AccountsReportingModel>> GetLedgerData(string StrAccNo, DateTime DTFrom, DateTime DTTo)
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                SELECT AccNo, AccTitle,VDate,VchrNo,Description,Debit,Credit,Balance  FROM VLedger
                WHERE AccNo = @StrAccNo
                AND VDate BETWEEN @DTFrom AND @DTTo
                ORDER BY VDate, SNo";

                var listdata = await db.QueryAsync<AccountsReportingModel>(sql,new { StrAccNo, DTFrom, DTTo });

                return listdata.ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching ledger data: {ex.Message}");
            throw;
        }
    }

    public async Task<decimal> GetAccountOpeningBalance(string strAccNo, DateTime OnDate)
    {
        string sql = "SELECT [dbo].[GetClosingBalance](@AccNo, @DT, @CurrentDT)";
        // Prepare the parameters
        var parameters = new
        {
            AccNo = strAccNo,
            DT = OnDate,
            CurrentDT = OnDate
        };
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            return await db.QueryFirstOrDefaultAsync<decimal>(sql, parameters);
        }
    }

    public async Task<List<AccountsReportingModel>> GetTransactionData(DateTime DTFrom, DateTime DTTo,string StrCond="")
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {

               var strcondition = "WHERE VDate BETWEEN @DTFrom AND @DTTo";
                    strcondition = strcondition + StrCond;

                strcondition = strcondition + " ORDER BY T1.VoucherType ASC,T1.VDate,T1.SNo";

                string sql = @"
                        SELECT VoucherType,AccNo,AccTitle,VDate,VchrNo,Description,Debit,Credit,BillDate,BillNo,
                               UserName AS GeneratedBy,MachineName,DT_Generator AS EntryDate,UserName_Edit AS EditedBy,
                               MachineName_Edit AS EditMachineName,DT_Edit EditedDate FROM VLedger1 T1 " + strcondition ;

                var listdata = await db.QueryAsync<AccountsReportingModel>(sql, new { DTFrom, DTTo , strcondition });
                return listdata.ToList();

            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching ledger data: {ex.Message}");
            throw;
        }
    }

    public async Task<List<AccountsReportingModel>> GetVoucherData(string StrVchrNo = "")
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {

                var strcondition = "WHERE VchrNo=@StrVchrNo";
                //strcondition = strcondition + StrVchrNo;

                string sql = @"
                        SELECT VoucherType,AccNo,AccTitle,VDate,VchrNo,Description,Debit,Credit,BillDate,BillNo,
                               UserName AS GeneratedBy,MachineName,DT_Generator AS EntryDate,UserName_Edit AS EditedBy,
                               MachineName_Edit AS EditMachineName,DT_Edit EditedDate FROM VLedger1 T1 " + strcondition;

                var listdata = await db.QueryAsync<AccountsReportingModel>(sql, new { StrVchrNo });
                return listdata.ToList();

            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching ledger data: {ex.Message}");
            throw;
        }
    }

    public async Task DeleteVoucher(AccountsReportingModel selectedVoucher, string strVchrNo, bool bDeleteForEdit=false)
    {
        using (IDbConnection db = new SqlConnection(_connectionString))
        {
            db.Open();
            using (var tran = db.BeginTransaction())
            {
                try
                {
                    string sqlBackupVoucher =
                        "INSERT INTO VouchersDeleted SELECT * FROM Vouchers WHERE VchrNo = @VchrNo";

                    string Username = selectedVoucher.UserName;
                    string MachineName = selectedVoucher.MachineName;
                    string Reason = selectedVoucher.DeleteReason;
                    string DT = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string sqlBackupInfo =
                        @"INSERT INTO VouchersDeletedInfo 
                      SELECT VchrNo, @Username, @MachineName, @DT, @Reason, UserName, MachineName,
                             DT,BillDate, BillNo 
                      FROM VoucherInfo 
                      WHERE VchrNo = @VchrNo";
                    
                    string sqlDeleteVoucher =
                        "DELETE FROM Vouchers WHERE VchrNo = @VchrNo";

                    string sqlDeleteVoucherInfo =
                        "DELETE FROM VoucherInfo WHERE VchrNo = @VchrNo";

                    if (!bDeleteForEdit)
                    {
                        await db.ExecuteAsync("DELETE FROM Advances WHERE AccVoucherNo = @VchrNo", new { VchrNo = strVchrNo }, tran);
                        await db.ExecuteAsync("DELETE FROM AmtCleared WHERE VchrNo = @VchrNo", new { VchrNo = strVchrNo }, tran);
                    }

                    await db.ExecuteAsync("DELETE FROM MakerPostedBills WHERE VchrNo = @VchrNo", new { VchrNo = strVchrNo }, tran);
                    await db.ExecuteAsync("DELETE FROM MakerAmtCleared WHERE VchrNo = @VchrNo", new { VchrNo = strVchrNo }, tran);
                    await db.ExecuteAsync("UPDATE VendIssued SET VchrNo = NULL WHERE VchrNo = @VchrNo", new { VchrNo = strVchrNo }, tran);
                    await db.ExecuteAsync("DELETE FROM MakerAdvancesShort WHERE AccVoucherNo = @VchrNo", new { VchrNo = strVchrNo }, tran);
                    await db.ExecuteAsync("DELETE FROM MakerAdvances WHERE AccVoucherNo = @VchrNo", new { VchrNo = strVchrNo }, tran);
                    await db.ExecuteAsync("DELETE FROM VenderPostedBills WHERE VchrNo = @VchrNo", new { VchrNo = strVchrNo }, tran);
                    await db.ExecuteAsync("DELETE FROM CustomInvoiceAuthorization WHERE VchrNo = @VchrNo", new { VchrNo = strVchrNo }, tran);

                    await db.ExecuteAsync(sqlBackupVoucher, new { VchrNo = strVchrNo }, tran);
                    await db.ExecuteAsync(sqlBackupInfo, new { VchrNo = strVchrNo, Username = Username, 
                          MachineName = MachineName, Reason= Reason, DT =DT}, tran);

                    await db.ExecuteAsync(sqlDeleteVoucher, new { VchrNo = strVchrNo }, tran);
                    await db.ExecuteAsync(sqlDeleteVoucherInfo, new { VchrNo = strVchrNo }, tran);

                    tran.Commit();
                }
                catch (Exception ex)
                {
                    tran.Rollback();
                    throw new Exception("Error On Deletion Of Voucher: " + ex.Message);
                }
            }
        }
    }
    public async Task<List<BOD_Meeting_Payment_List_ViewModel>> GetBODMeetingPaymentList(DateTime DTFrom, DateTime DTTo, int iStatus=3,int iBankID=0)
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {

                var strcondition = " WHERE DT BETWEEN @DTFrom AND @DTTo";
                if (iStatus == 0) //Under Process
                    strcondition += " AND ISNULL(Status,0)=0 AND ISNULL(PostedVchrNo,'')=''";
                else if (iStatus==1)    //Approved & Posted
                    strcondition += " AND ISNULL(Status,0)=1 AND ISNULL(PostedVchrNo,'')<>''";
                else if (iStatus == 2)    //Voucher Posted
                    strcondition += " AND ISNULL(Status,0)=0 AND ISNULL(PostedVchrNo,'')<>''";

                if (iBankID != 0)
                    strcondition += " AND BankID=@iBankID";

                string sql = @"
                        SELECT EntryID,DT,SIAL_Bank_Acc_No,SIAL_Bank_Acc_Title,Debit_Acc_No,Debit_Acc_Title,Tax_Acc_No,Tax_Acc_Title
                               ,TotalDebit,TotalTax,TotalDebit,BankID,TotalEntries,PostedVchrNo,Status,Transaction_Type,Remarks
                               FROM VOnline_PV T1 " + strcondition;

                //var listdata = await db.QueryAsync<BOD_Meeting_Payment_List_ViewModel>(sql, new { StrVchrNo });
                var listdata = await db.QueryAsync<BOD_Meeting_Payment_List_ViewModel>(sql,new {DTFrom,DTTo,iBankID });
                return listdata.ToList();

            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching ledger data: {ex.Message}");
            throw;
        }
    }
    public async Task<List<BOD_Meeting_Payment_Detail_ViewModel>> GetBODMeetingPaymentDetail(int EntryID)
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {

                var strcondition = " WHERE OL_PV_RefID=@EntryID";
                
                string sql = @"
                        SELECT EntryID,OL_PV_RefID,SD_RefID,Bank_Acc_No,Bank_Code,Bank_Acc_Title,Bank_Name,Debit,Tax,Status,SMS_Status
                               ,FolioNo,DirectorName,BOD_Meetings_RefID,ChqNo,Advising,Transaction_Type,BankID,Bank_Accurate_Name
                               ,Bank_Code_Name,VchrNo FROM VOnline_PV_Detail T1 " + strcondition;

                //var listdata = await db.QueryAsync<BOD_Meeting_Payment_List_ViewModel>(sql, new { StrVchrNo });
                var listdata = await db.QueryAsync<BOD_Meeting_Payment_Detail_ViewModel>(sql, new { EntryID });
                return listdata.ToList();

            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching GetBODMeetingPaymentDetail data: {ex.Message}");
            throw;
        }
    }
    public async Task<List<Online_Payment_Voucher_List_ViewModel>> GetOnlinePaymentVoucherList(DateTime DTFrom, DateTime DTTo, int iBankID = 0, int iStatus = 0,int iVchrType=0)
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {

                var strcondition = " WHERE VDate BETWEEN @DTFrom AND @DTTo";
                if (iBankID != 0)
                    strcondition += " AND BankID=@iBankID";

                if (iStatus == 1)
                    strcondition += " AND Online_Posting=false";
                else if(iStatus == 2)
                    strcondition += " AND Online_Posting=true";

                if (iVchrType==1)
                    strcondition += " AND Online_Vchr_Type=2";
                else if (iVchrType == 2)
                    strcondition += " AND Online_Vchr_Type=1";

                string sql = @"
                        SELECT VchrNo,VDate,Payee,ChequeNo,Amount,Advising,Online_Posting,M1No,BankID,Bank,Branch,City,AccNo,Online_Vchr_Type"+
                               " FROM VVouchers_Online_List T1 " + strcondition + " ORDER BY T1.BankID,T1.VDate,T1.ChequeNo";
                
                var listdata = await db.QueryAsync<Online_Payment_Voucher_List_ViewModel>(sql, new { DTFrom, DTTo, iBankID });
                return listdata.ToList();

            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching ledger data: {ex.Message}");
            throw;
        }
    }
    public async Task<List<Trial_Balance_ViewModel>> GetTrialBalanceList(DateTime DTFrom, DateTime DTTo,string Code)
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var parameters = new
                {
                    DTFrom,
                    DTTo,
                    Code                    
                };

                var listdata = await db.QueryAsync<Trial_Balance_ViewModel>(
                "Trial_Balance_SP",          // your stored procedure name
                parameters,
                commandType: CommandType.StoredProcedure);

                return listdata.ToList();                

            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching trial balance data: {ex.Message}");
            throw;
        }
    }
    public async Task<List<Bank_Balance_Statement_ViewModel>> GetBankBalanceStatement()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {

                var strcondition = " ";

                string sql = @"
                        SELECT AccNo,Bank,Branch,BankAccNo,Balance,ChqFormat FROM VBankAccounts";

                var listdata = await db.QueryAsync<Bank_Balance_Statement_ViewModel>(sql);
                return listdata.ToList();

            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching Bank Balance Statement Data: {ex.Message}");
            throw;
        }
    }
    public async Task<List<Cheque_Receiving_ViewModel>> GetChequeReceivingData(DateTime DTFrom, DateTime DTTo)
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var strcondition = " WHERE DT BETWEEN @DTFrom AND @DTTo";
                
                string sql = @"
                        SELECT EntryID,DT,PartyName,ChqNo,DepositBank,Amount,DeptName,Name,EmpID,ReceivedFrom,DTChq,SlipNo,DTSlip" +
                               " FROM VChequeReceivingReport T1 " + strcondition + " ORDER BY T1.DT";

                var listdata = await db.QueryAsync<Cheque_Receiving_ViewModel>(sql, new { DTFrom, DTTo });
                return listdata.ToList();

            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching Cheque Receiving: {ex.Message}");
            throw;
        }
    }
    public async Task<List<Important_Tasks_ViewModel>> GetImportantTasksData(DateTime DTFrom, DateTime DTTo)
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var strcondition = " WHERE DT BETWEEN @DTFrom AND @DTTo";

                string sql = @"
                        SELECT EntryID,DT,Title,Particulars,DutyAssignedTo,Status,DTDue,OK" +
                               " FROM ImportantTasks T1 " + strcondition + " ORDER BY T1.DT";

                var listdata = await db.QueryAsync<Important_Tasks_ViewModel>(sql, new { DTFrom, DTTo });
                return listdata.ToList();

            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching Important Tasks: {ex.Message}");
            throw;
        }
    }
    public async Task<List<Bills_Received_ViewModel>> GetBillsReceivedData(DateTime DTFrom, DateTime DTTo,int SerialFrom,int SerialTo)
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string strcondition = "";
                if (SerialFrom>0)
                    strcondition = " WHERE EntryID BETWEEN @SerialFrom AND @SerialTo";
                else
                    strcondition = " WHERE DTIn BETWEEN @DTFrom AND @DTTo";

                string sql = @"
                        SELECT EntryID,DTIn,RefNo,Party,Description,ReceivedFrom,ReceivedBy,HandedOverTo,Status,Objection,Amount,DTDue,DTHandedOver,Remarks" +
                               " FROM BillsReceived T1 " + strcondition + " ORDER BY T1.DTIn";

                var listdata = await db.QueryAsync<Bills_Received_ViewModel>(sql, new { DTFrom, DTTo, SerialFrom, SerialTo });
                return listdata.ToList();

            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching Bills Received Data: {ex.Message}");
            throw;
        }
    }
    public async Task<List<NTN_CNIC_ViewModel>> GetNTN_CNICData(string SearchText)
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string strcondition = "";
                if (!SearchText.Equals(string.Empty))
                    strcondition = $" WHERE (Name LIKE @SearchText) OR (NTNNo LIKE @SearchText) OR (CNICNo LIKE @SearchText)";

                string sql = @"
                        SELECT EntryID,ReferenceNo,Name,NTNNo,CNICNo" +
                               " FROM NTN_CNIC_Nos T1 " + strcondition + " ORDER BY T1.EntryID";

                var listdata = await db.QueryAsync<NTN_CNIC_ViewModel>(sql,new { SearchText = $"%{SearchText}%" });
                return listdata.ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching NTN CNIC Data: {ex.Message}");
            throw;
        }
    }
    public async Task<List<Letter_Dispatch_Register_ViewModel>> GetLetterDispatchData(DateTime DTFrom, DateTime DTTo)
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var strcondition = " WHERE DT BETWEEN @DTFrom AND @DTTo";

                string sql = @"
                        SELECT EntryID,DT,RefNo,PartyName,Address,CellNo,CourierName,CourierNo" +
                               " FROM LetterDispatchRegister T1 " + strcondition + " ORDER BY T1.DT";

                var listdata = await db.QueryAsync<Letter_Dispatch_Register_ViewModel>(sql, new { DTFrom, DTTo });
                return listdata.ToList();

            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching Leeter Dispatch Register: {ex.Message}");
            throw;
        }
    }
}
