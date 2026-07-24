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
                SELECT V.SNo, V.AccNo, V.AccTitle, V.VDate, V.VchrNo, V.Description, V.Debit, V.Credit, V.Balance, 
                       V.chequeno AS ChqNo, V.chequeDate AS ChqDate, V.chequeType AS ChqType, 
                       V.Payee, V.Handed_Over_To, V.UserName, V.MachineName, V.DT AS EntryDate,
                       B.UserName AS BalanceTag_UserName, B.DTEntry AS BalanceTag_DTEntry
                FROM VLedger V
                LEFT JOIN BalanceTags B ON V.SNo = B.Sno
                WHERE V.AccNo = @StrAccNo
                AND V.VDate BETWEEN @DTFrom AND @DTTo
                ORDER BY V.VDate, V.SNo";

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

    public async Task<ChequeDetailModel> GetChequeDetails(string vchrNo, string chqNo)
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                SELECT CDate, chequeno AS ChequeNo, Bank, Branch, Description, Amount, 
                       chequeType AS ChequeType, chequeDate AS ChequeDate, Posted, 
                       Bounced, ClearanceDT, Payment
                FROM VChqLedger
                WHERE VchrNo = @vchrNo AND chequeno = @chqNo";

                return await db.QueryFirstOrDefaultAsync<ChequeDetailModel>(sql, new { vchrNo, chqNo });
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching cheque details: {ex.Message}");
            return null;
        }
    }

    public async Task InsertBalanceTag(double sNo, string userName)
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = "INSERT INTO BalanceTags(Sno, UserName) VALUES(@sNo, @userName)";
                await db.ExecuteAsync(sql, new { sNo, userName });
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error inserting balance tag: {ex.Message}");
            throw;
        }
    }

    public async Task RemoveBalanceTag(double sNo)
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = "DELETE FROM BalanceTags WHERE Sno = @sNo";
                await db.ExecuteAsync(sql, new { sNo });
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error removing balance tag: {ex.Message}");
            throw;
        }
    }

    public async Task<decimal> GetAccountOpeningBalance(string strAccNo, DateTime OnDate)
    {
        string sql = "SELECT [dbo].[GetBalance](@AccNo, @DT, @CurrentDT)";
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

                strcondition = strcondition + " ORDER BY VoucherType ASC, T1.VDate, T1.SNo";

                string sql = @"
                        SELECT 
                            CASE
                                WHEN LEFT(VchrNo, 3) = 'BPV' THEN 'BPV'
                                WHEN LEFT(VchrNo, 3) = 'CPV' THEN 'CPV'
                                WHEN LEFT(VchrNo, 3) = 'BRV' THEN 'BRV'
                                WHEN LEFT(VchrNo, 3) = 'CRV' THEN 'CRV'
                                ELSE 'JV'
                            END AS VoucherType,
                            AccNo, AccTitle, VDate, VchrNo, Description, Debit, Credit, 
                            UserName AS GeneratedBy, MachineName, DT AS EntryDate, Handed_Over_To, chequeno AS ChqNo, chequeDate AS ChqDate 
                        FROM VLedger T1 " + strcondition;

                var listdata = await db.QueryAsync<AccountsReportingModel>(sql, new { DTFrom, DTTo });
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
                        SELECT 
                            CASE
                                WHEN LEFT(VchrNo, 3) = 'BPV' THEN 'BPV'
                                WHEN LEFT(VchrNo, 3) = 'CPV' THEN 'CPV'
                                WHEN LEFT(VchrNo, 3) = 'BRV' THEN 'BRV'
                                WHEN LEFT(VchrNo, 3) = 'CRV' THEN 'CRV'
                                ELSE 'JV'
                            END AS VoucherType,
                            AccNo, AccTitle, VDate, VchrNo, Description, Debit, Credit, 
                            UserName AS GeneratedBy, MachineName, VchrEntryDate AS EntryDate 
                        FROM VLedger1 T1 " + strcondition;

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
                commandTimeout: 180,
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
    public async Task<List<Cash_Balance_Statement_ViewModel>> GetCashBalanceStatement()
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string cashInHandParent = await db.QueryFirstOrDefaultAsync<string>("SELECT DataValue FROM GeneralData WHERE DataName = 'CashInHandParent'");
                if (string.IsNullOrEmpty(cashInHandParent))
                {
                    cashInHandParent = "15-001";
                }

                string sql = "SELECT AccNo, AccTitle, Balance FROM VActiveAccounts WHERE LEFT(AccNo, 6) = @cashInHandParent";
                var listdata = await db.QueryAsync<Cash_Balance_Statement_ViewModel>(sql, new { cashInHandParent });
                return listdata.ToList();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching Cash Balance Statement Data: {ex.Message}");
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

    public async Task<List<Cash_Book_Report_ViewModel>> GetCashBookReport(DateTime dtFrom, DateTime dtTo)
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var parameters = new
                {
                    DTFrom = dtFrom,
                    DTTo = dtTo
                };

                var listdata = await db.QueryAsync<Cash_Book_Report_ViewModel>(
                    "CashBookReport_SP",
                    parameters,
                commandTimeout: 180,
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

    public async Task PrepareExpenseReportDataAsync(DateTime dtStart)
    {
        try
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                DateTime dtEnd = dtStart.AddMonths(6).AddDays(-1);
                DateTime myDtEnd = dtStart.AddMonths(12).AddDays(-1);
                DateTime dtPrevStart = dtStart.AddYears(-1);
                DateTime dtPrevEnd = myDtEnd.AddYears(-1);

                db.Open();
                using (var transaction = db.BeginTransaction())
                {
                    await db.ExecuteAsync("DELETE FROM rptSaleHeaders", transaction: transaction);
                    await db.ExecuteAsync("DELETE FROM rptSalePeriods", transaction: transaction);
                    await db.ExecuteAsync("DELETE FROM rptSalePeriodsPrevious", transaction: transaction);

                    // Build Periods list
                    var periods = new List<string>();
                    DateTime current = dtStart;
                    while (current <= myDtEnd)
                    {
                        string p = current.ToString("MMM-yy");
                        periods.Add(p);
                        await db.ExecuteAsync("INSERT INTO rptSalePeriods(Period) VALUES(@p)", new { p }, transaction);
                        current = current.AddMonths(1);
                    }

                    while (periods.Count < 12) periods.Add("");

                    string insertHeader = $@"
                        INSERT INTO rptSaleHeaders(Period1, Period2, Period3, Period4, Period5, Period6, Period7, Period8, Period9, Period10, Period11, Period12)
                        VALUES('{periods[0]}', '{periods[1]}', '{periods[2]}', '{periods[3]}', '{periods[4]}', '{periods[5]}', 
                               '{periods[6]}', '{periods[7]}', '{periods[8]}', '{periods[9]}', '{periods[10]}', '{periods[11]}')";
                    await db.ExecuteAsync(insertHeader, transaction: transaction);

                    current = dtPrevStart;
                    while (current <= dtPrevEnd)
                    {
                        string p = current.ToString("MMM-yy");
                        await db.ExecuteAsync("INSERT INTO rptSalePeriodsPrevious(Period) VALUES(@p)", new { p }, transaction);
                        current = current.AddMonths(1);
                    }

                    long lRefID = await db.ExecuteScalarAsync<long>("SELECT ISNULL(MAX(EntryID), 1) FROM rptSaleHeaders", transaction: transaction);

                    string pivotInsertSales = @"
                        INSERT INTO rptSales (RefID, Customer, Period1, Period2, Period3, Period4, Period5, Period6, Period7, Period8, Period9, Period10, Period11, Period12)
                        SELECT 
                            @lRefID,
                            AccNo,
                            ISNULL([1],0) AS Period1, ISNULL([2],0) AS Period2, ISNULL([3],0) AS Period3, ISNULL([4],0) AS Period4, 
                            ISNULL([5],0) AS Period5, ISNULL([6],0) AS Period6, ISNULL([7],0) AS Period7, ISNULL([8],0) AS Period8, 
                            ISNULL([9],0) AS Period9, ISNULL([10],0) AS Period10, ISNULL([11],0) AS Period11, ISNULL([12],0) AS Period12
                        FROM (
                            SELECT AccNo, Amount, ROW_NUMBER() OVER(PARTITION BY AccNo ORDER BY SortID) as PeriodNum 
                            FROM dbo.F_Expenses(@DTStart, @DTEnd)
                        ) as SourceTable
                        PIVOT (
                            SUM(Amount)
                            FOR PeriodNum IN ([1], [2], [3], [4], [5], [6], [7], [8], [9], [10], [11], [12])
                        ) AS PivotTable
                        WHERE TRY_CAST(Left(AccNo,2) AS INT) >= 31";

                    await db.ExecuteAsync(pivotInsertSales, new { lRefID, DTStart = dtStart, DTEnd = dtEnd }, transaction: transaction);

                    string pivotInsertSalesPrevious = @"
                        INSERT INTO rptSalesPrevious (RefID, Customer, Period1, Period2, Period3, Period4, Period5, Period6, Period7, Period8, Period9, Period10, Period11, Period12)
                        SELECT 
                            @lRefID,
                            AccNo,
                            ISNULL([1],0) AS Period1, ISNULL([2],0) AS Period2, ISNULL([3],0) AS Period3, ISNULL([4],0) AS Period4, 
                            ISNULL([5],0) AS Period5, ISNULL([6],0) AS Period6, ISNULL([7],0) AS Period7, ISNULL([8],0) AS Period8, 
                            ISNULL([9],0) AS Period9, ISNULL([10],0) AS Period10, ISNULL([11],0) AS Period11, ISNULL([12],0) AS Period12
                        FROM (
                            SELECT AccNo, Amount, ROW_NUMBER() OVER(PARTITION BY AccNo ORDER BY SortID) as PeriodNum 
                            FROM dbo.F_ExpensesPrevious(@DTPrevStart, @DTPrevEnd)
                        ) as SourceTable
                        PIVOT (
                            SUM(Amount)
                            FOR PeriodNum IN ([1], [2], [3], [4], [5], [6], [7], [8], [9], [10], [11], [12])
                        ) AS PivotTable
                        WHERE TRY_CAST(Left(AccNo,2) AS INT) >= 31";

                    await db.ExecuteAsync(pivotInsertSalesPrevious, new { lRefID, DTPrevStart = dtPrevStart, DTPrevEnd = dtPrevEnd }, transaction: transaction);

                    transaction.Commit();
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error preparing expense report data: {ex.Message}");
            throw;
        }
    }
}
