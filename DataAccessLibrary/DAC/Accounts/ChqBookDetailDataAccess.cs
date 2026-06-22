using Dapper;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using OfficeOpenXml.Drawing.Chart;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Security.Cryptography.X509Certificates;
using System.Security.Principal;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace DataAccessLibrary.DAC.Accounts
{
    public class ChqBookDetailDataAccess : IChqBookDetailDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<VouchersDataAccess> _logger;
        public ChqBookDetailDataAccess(IConfiguration configuration, ILogger<VouchersDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection");
            _logger = logger;
        }

        public async Task<List<GenericDropDownModel>> GetValuesForDropDown(string strTableName, string strFieldName_ID, string strFieldName_Description, string strCondition = "")
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {

                    string sql = @"SELECT CAST(" + strFieldName_ID + " AS VARCHAR(50)) AS DropDownValue_ID," + strFieldName_Description + " AS DropDownValue_Description FROM " + strTableName + " " + strCondition;
                    return (await db.QueryAsync<GenericDropDownModel>(sql)).ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching ChartOfAccounts: {ex.Message}");
                throw;
            }
        }

        public async Task<List<BankAccountInfo>> GetBankAccounts()
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = @"SELECT BankID,AccNo,AccTitle FROM VBankAccounts ORDER BY VBankAccounts.AccTitle";
                    return (await db.QueryAsync<BankAccountInfo>(sql)).ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching ChartOfAccounts: {ex.Message}");
                throw;
            }
        }

        public async Task<string> ExecuteChequeBookSave(ChqBookDetailViewModel chqbook)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                db.Open();
                using (var transaction = db.BeginTransaction())
                {
                    try
                    {
                        long nextchqBookNo = await db.QuerySingleAsync<long>(
                            "dbo.InsertChqBookData_SP",
                            new
                            {
                                StartingFrom = chqbook.StartingFrom,
                                Chqs = chqbook.Chqs,
                                ManualNo = chqbook.ManualNo,
                                ChqBookDetail = chqbook.ChqBookDetail,
                                AccNo = chqbook.AccNo,
                                UserName = chqbook.UserName,
                                MachineName = chqbook.MachineName,
                                Chq_Type = chqbook.Chq_Type
                            },
                            commandType: CommandType.StoredProcedure, transaction: transaction
                        );

                        long? returnSNo = null;
                        chqbook.ChqBookNo = nextchqBookNo;
                        returnSNo = await InsertChequeDetailsAsync(db, transaction, chqbook, nextchqBookNo);
                        transaction.Commit();
                        return nextchqBookNo.ToString();

                    }
                    catch (Exception ex)
                    {
                        {
                            transaction.Rollback();
                            _logger.LogError(
                                ex,
                                "Error: {ErrorMessage}",
                                chqbook.StartingFrom,
                                chqbook.Chqs,
                                ex.Message
                            );
                            throw;
                        }
                    }
                }
            }
        }

        private async Task<long?> InsertChequeDetailsAsync(IDbConnection db,IDbTransaction transaction,ChqBookDetailViewModel chequeDetails, long ChqBookNo)
        {
            if (chequeDetails == null)
                return null;

                long lastChequeNo = await db.QuerySingleAsync<long>(
                    "dbo.InsertChqBookDetailData_SP",
                    new
                    {
                        StartingFrom = chequeDetails.StartingFrom,
                        Chqs = chequeDetails.Chqs,
                        ChqBookNo = chequeDetails.ChqBookNo
                    },
                    commandType: CommandType.StoredProcedure,
                    transaction: transaction
                );

                return lastChequeNo;
        }

        public async Task<List<ChqBookDetailViewModel>> GetBankChqBooks(string StrAccNo="")
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = @"SELECT ChqBookNo,AccNo,StartingFrom,Chqs,ManualNo,ChqBookDetail,Chq_Type,ChqsLeft 
                        FROM VChqBooks WHERE @StrAccNo='' OR AccNo=@StrAccNo";
                    return (await db.QueryAsync<ChqBookDetailViewModel>(sql, new { StrAccNo=StrAccNo ?? "" })).ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching ChqBookDetail: {ex.Message}");
                throw;
            }
        }

        public async Task<bool> DeleteChqBook(long ChqBookNo)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = @"DELETE FROM ChqBooks WHERE ChqBookNo = @ChqBookNo";
                    return await db.ExecuteAsync(sql, new { ChqBookNo }) > 0;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                throw;
            }
        }
        public async Task<List<BankAccountInfo>> GetBankAccountsList()
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = @"SELECT BankID,AccNo,AccTitle,Bank,Branch,City,BankAccNo,DefaultBank,AllowOD,ODAmt,ODRate
                                 FROM VBankAccounts ORDER BY DefaultBank DESC,Bank ASC";
                    return (await db.QueryAsync<BankAccountInfo>(sql)).ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching ChqBookDetail: {ex.Message}");
                throw;
            }
        }

        public async Task<bool> SaveOverDraftData(BankAccountInfo bankAccount, string strAccno,bool bdefaultaccount)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {

                    string sql = "";
                    if (bdefaultaccount == true)
                    {

                        string sqlall = @"Update BankAccounts Set DefaultBank=0";
                        await db.ExecuteAsync(sqlall);

                        sql = @"Update BankAccounts Set DefaultBank=@DefaultBank WHERE AccNo= @Accno";
                        var para = new
                        {
                            Accno = strAccno,
                            DefaultBank = bdefaultaccount
                        };
                        return await db.ExecuteAsync(sql, para) > 0;
                    }
                    else
                    {
                        sql = @"Update BankAccounts Set AllowOD=@AllowOD,ODAmt=@ODAmt,ODRate=@ODRate,DefaultBank=@DefaultBank WHERE AccNo= @Accno";
                        var parameters = new
                        {
                            Accno = strAccno,
                            AllowOD = bankAccount.AllowOD,
                            ODAmt = bankAccount.ODAmt,
                            ODRate = bankAccount.ODRate,
                            DefaultBank = bankAccount.DefaultBank
                        };
                        return await db.ExecuteAsync(sql, parameters) > 0;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                throw;
            }
        }

        public async Task<List<ChqListModel>> GetChqBookChequesList(long chqbookno)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = @"SELECT ChqBookNo,ChqNo,ChequeType AS Chqtype,CDate,ChequeDate,AccNo,AccTitle,Description,Amount,Payee,ClearanceDT,Issued,Bounced,DeleteDescription,sno AS CSNo,Payee AS chqpayee FROM VChqBookLedger WHERE ChqBookNo=@chqbookno Order By Cast(ChqNo as numeric)";
                    return (await db.QueryAsync<ChqListModel>(sql, new { chqbookno = chqbookno })).ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching ChqBookDetail: {ex.Message}");
                throw;
            }
        }
        public async Task<bool> ChangeChqStatus(ChqListModel chqs, string chqno,double sno, int iType,long INextSno)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = "";
                    if (iType == 0)  // MAKE CHQ UN CANCEL
                    {
                        string strdel = @"DELETE FROM [Cheque] WHERE Sno = @sno";
                        await db.ExecuteAsync(strdel, new { sno });

                        string sqlall =@"UPDATE [ChqList] SET issued = 0 WHERE chqno = @chqno";
                        return await db.ExecuteAsync(sqlall, new { chqno }) > 0;
                    }
                    else if (iType == 1) // MAKE CHQ CANCEL
                    {
                        //long INextSno = 0;// INextSno(Date, "Cheque")
                        sql = @"INSERT INTO Cheque(SNo,CDate,AccNo,Description,Amount,chequeno,ChqBookNo,chequeType,chequeDate,Posted,Payment,Bounced) VALUES 
                                (@INextSno,@Date,@AccNo,@Reason,0,@ChqNo,@ChqBookNo,'Canceled',@Date,1,1,0)";
                        var parameters = new
                        {
                            INextSno = @INextSno,
                            Date = DateTime.Now.Date ,
                            AccNo = chqs.AccNo,
                            ChqNo= chqs.ChqNo,
                            ChqBookNo=chqs.ChqBookNo,
                            Reason=chqs.CancelReason
                            
                        };
                        await IssueCheque(chqs.ChqBookNo, chqs.ChqNo);
                        return await db.ExecuteAsync(sql, parameters) > 0;
                    }
                    else if (iType == 2) // CHANGE PRINTING STATE
                    {

                        sql = "INSERT INTO ChequePrintHistory(Sno,Payee,PrintDT,Designation,Company,UserName,MachineName) VALUES" +
                            "(@sno,@payee,@chqprintdate,@chqdesignation,@chqcompany,@username,@machinename)";
                        var parameters = new
                        {
                            sno = @sno,
                            payee = chqs.chqpayee,
                            chqdesignation = chqs.chqdesignation,
                            chqcompany = chqs.chqcompany,
                            chqprintdate = chqs.chqprintdate,
                            Reason = chqs.CancelReason,
                            username=chqs.UserName,
                            machinename = chqs.MachineName
                        };
                        await db.ExecuteAsync(sql, parameters );

                        sql = "";
                        sql = @"UPDATE Cheque SET ChqPrintingDone=1,ChqDesignation=@chqdesignation,
                                ChqCompany=@chqcompany,ChqPrintDate=@chqprintdate
                                WHERE SNo=@sno";
                        var para = new
                        {
                            sno= @sno,
                            chqdesignation = chqs.chqdesignation,
                            chqcompany = chqs.chqcompany,
                            chqprintdate = chqs.chqprintdate,
                        };
                        return await db.ExecuteAsync(sql, para ) > 0;
                    }
                    else
                    {
                        return false;
                    }                        
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                throw;
            }
        }

        public async Task<bool> IssueCheque(long ChqBookNo,string Chqno)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"UPDATE ChqList SET Issued=1 WHERE ChqNo= @Chqno AND ChqBookNo=@ChqBookNo";
                return await db.ExecuteAsync(sql, new { Chqno, ChqBookNo }) > 0;
            }
        }

    }
}


