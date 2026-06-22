using Dapper;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics.Metrics;
using static Azure.Core.HttpHeader;

namespace DataAccessLibrary.DAC.Accounts
{
    public class VouchersDataAccess : IVouchersDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<VouchersDataAccess> _logger;
        private readonly IDBHelper _db;
        public VouchersDataAccess(IConfiguration configuration, ILogger<VouchersDataAccess> logger, IDBHelper db)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                            ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");

            _logger = logger;
            _db = db;
        }

        public async Task<List<GenericDropDownModel>> GetValuesForDropDown(string strTableName,string strFieldName_ID,string strFieldName_Description,string strCondition="")
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {

                    string sql = @"SELECT CAST("+strFieldName_ID+" AS VARCHAR(50)) AS DropDownValue_ID,"+strFieldName_Description+" AS DropDownValue_Description FROM "+strTableName+" "+strCondition;                    
                    return (await db.QueryAsync<GenericDropDownModel>(sql)).ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching ChartOfAccounts: {ex.Message}");
                throw; // Rethrow or handle appropriately
            }
        }
        public async Task<List<SIALDirectorViewModel>> GetDirectorsByBODMeetingDate(DateTime meetingDate)
        {
            int iBOD_EntryID = _db.getSingleIntValue("EntryID", "BOD_Meetings", " WHERE DT=@MeetingDate",
                new { MeetingDate = meetingDate }
                );

            if (iBOD_EntryID == 0) return new List<SIALDirectorViewModel>();

            string strCondition = " WHERE ISNULL(FolioNo,'')<>'' AND ISNULL(DirectorName,'')<>''";
            strCondition += " AND EntryID IN(SELECT SD_RefID FROM BOD_Meetings_Participants WHERE BOD_Meetings_RefID=@BOD_EntryID)";
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {

                    string sql = @"SELECT EntryID,FolioNo,DirectorName,Company,Bank_Name,Bank_Account_Title,Bank_Account_No,Bank_Code,T1.SD_RefID FROM SIALDirectors "+
                                    "LEFT JOIN (SELECT SD_RefID FROM VOnline_PV_Detail WHERE BOD_Meetings_RefID=@BOD_EntryID) T1 ON SIALDirectors.EntryID=T1.SD_RefID " + 
                                    strCondition+ " ORDER BY T1.SD_RefID,SIALDirectors.Bank_Account_Title DESC";

                    return (await db.QueryAsync<SIALDirectorViewModel>(sql,new {BOD_EntryID=iBOD_EntryID })).ToList();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching Directors for BOD Meeting ID: {MeetingID}", iBOD_EntryID);                
                throw; // Rethrow or handle appropriately
            }
        }
        public async Task<List<GenericDropDownModel>> GetBPVPaymentTypes()
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {

                    string sql = @"SELECT EnumValue AS DropDownValue_ID,EnumDescription AS DropDownValue_Description FROM EnumValues WHERE EnumName='BPV_Payment_Types' ORDER BY EnumValue";

                    /*var paymentMethods = await db.QueryAsync<GenericDropDownModel>(sql, new { });
                    return paymentMethods.ToList();*/

                    return (await db.QueryAsync<GenericDropDownModel>(sql)).ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching ChartOfAccounts: {ex.Message}");
                throw; // Rethrow or handle appropriately
            }
        }
        public async Task<List<GenericDropDownModel>> GetOnlineVoucherPaymentModes()
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {

                    string sql = @"SELECT EnumValue AS DropDownValue_ID,EnumDescription AS DropDownValue_Description FROM EnumValues WHERE EnumName='Online_Voucher_Payment_Modes' ORDER BY EnumValue";
                 
                    return (await db.QueryAsync<GenericDropDownModel>(sql)).ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching ChartOfAccounts: {ex.Message}");
                throw; 
            }
        }
        public async Task<List<ChartOfAccountsModel>> GetTransactionalAccounts(string? strCondition = null)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    string sql = @"SELECT AccNo,AccTitle FROM VActiveAccounts";
                    if (!string.IsNullOrWhiteSpace(strCondition))
                    {
                        sql += " "+strCondition;
                    }
                    sql += " ORDER BY VActiveAccounts.AccTitle";
                    return (await db.QueryAsync<ChartOfAccountsModel>(sql)).ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching ChartOfAccounts: {ex.Message}");
                throw; // Rethrow or handle appropriately
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
                throw; // Rethrow or handle appropriately
            }
        }
        
        public async Task<ChequeInfo> GetNewChequeNumbersByBankAccount(string strAccNo)
        {
            string sql = @"SELECT TOP 1 ChqBookNo,ChqNo,Issued FROM VChqList WHERE AccNo=@strAccNo AND Issued=0 ORDER BY CAST(ChqNo AS NUMERIC)";
            
            using (IDbConnection db = new SqlConnection(_connectionString))
            {                
                return await db.QueryFirstOrDefaultAsync<ChequeInfo>(sql, new { @strAccNo = strAccNo });
            }
        }
        public async Task<decimal> GetBalance(string strAccNo, DateTime OnDate)
        {
            string sql = "SELECT [dbo].[GetBalance](@AccNo, @DT, @CurrentDT)";
            // Prepare the parameters
            var parameters = new
            {
                AccNo = strAccNo,
                DT = OnDate,
                CurrentDT = DateTime.Now
            };
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                return await db.QueryFirstOrDefaultAsync<decimal>(sql, parameters);
            }
        }
        public async Task<string> GetNextVchrNo(DateTime OnDate,string VoucherType)
        {
            string sql = "SELECT [dbo].[GetNextVchrNo](@DT,@VchrType)";
            // Prepare the parameters
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
        public async Task<string> GetNextVchrNo_Approval(DateTime OnDate, string VoucherType)
        {
            string sql = "SELECT [dbo].[GetNextVchrNo_Approval](@DT,@VchrType)";            
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
        public async Task<long> GetNextSNo(DateTime OnDate,string strTableName,IDbConnection? existingDb = null,IDbTransaction? existingTrans = null)
        {
            string sql = "";
            if (strTableName=="Cheque")
                sql="SELECT MAX(SNo) AS MaxSNo FROM Cheque WHERE CDate=@VDate";
            else
                sql = "SELECT MAX(SNo) AS MaxSNo FROM Vouchers WHERE VDate=@VDate";
            // Prepare the parameters
            var parameters = new
                {
                    VDate = OnDate
                };
            bool isInternalConn = (existingDb == null);
            IDbConnection db = existingDb ?? new SqlConnection(_connectionString);
            try
            {
                if (db.State != ConnectionState.Open) db.Open();
                
                long? maxSNo = await db.QueryFirstOrDefaultAsync<long?>(sql, parameters, transaction: existingTrans);
                long result = 0;
                if (maxSNo == null)
                {
                    string formattedString = OnDate.ToString("yyMMdd") + "0001";
                    result = long.Parse(formattedString);
                }
                else
                    result = (maxSNo ?? 1) + 1;

                return result;                
            }
            finally
            {
                // 3. ONLY close the connection if we created it here
                if (isInternalConn) db.Dispose();
            }

        }
        public async Task<string> GetNextChqNo(string accNo,double chqBookNo = 0,double lastChqNo = 0,int iChqType = 0,IDbConnection? existingDb = null,IDbTransaction? existingTrans = null)
        {
            // 1. Build the condition
            // We use WITH (UPDLOCK) to prevent another user from grabbing this same cheque number 
            // before we finish our transaction.
            string sql = @"SELECT TOP 1 ChqNo FROM VChqList WITH (UPDLOCK, HOLDLOCK) WHERE AccNo = @AccNo AND CAST(ChqNo AS NUMERIC) > @LastChq AND Issued = 0 AND Chq_Type = @ChqType";

            if (chqBookNo > 0)
            {
                sql += " AND ChqBookNo = @BookNo";
            }

            sql += " ORDER BY CAST(ChqNo AS NUMERIC) ASC";

            var parameters = new
            {
                AccNo = accNo,
                LastChq = lastChqNo,
                ChqType = iChqType,
                BookNo = chqBookNo
            };

            // 2. Manage Connection
            bool isInternalConn = (existingDb == null);
            IDbConnection db = existingDb ?? new SqlConnection(_connectionString);

            try
            {
                if (db.State != ConnectionState.Open) db.Open();

                // 3. Execute Query
                var result = await db.QueryFirstOrDefaultAsync<string>(sql, parameters, transaction: existingTrans);

                // Return "0" to match your VB6 logic if nothing is found
                return result ?? "0";
            }
            finally
            {
                if (isInternalConn) db.Dispose();
            }
        }
        public async Task<string> ExecuteVoucherSave(
            VoucherViewModel voucher,
            IDbConnection? existingDb = null,
            IDbTransaction? existingTrans = null)
        {
            bool isInternalConn = (existingDb == null);
            IDbConnection db = existingDb ?? new SqlConnection(_connectionString);
            if (db.State != ConnectionState.Open) db.Open();
            IDbTransaction transaction = existingTrans ?? db.BeginTransaction();
            try
            {
                string strVchrNo = await db.QuerySingleAsync<string>(
                    "dbo.InsertVoucherHead_SP",
                    new
                    {
                        VchrNo = voucher.VchrNo,
                        UserName = voucher.UserName,
                        MachineName = voucher.MachineName,
                        DT = voucher.DT,
                        Notes = voucher.Notes,
                        Payee = voucher.Payee,
                        TaxType = voucher.TaxType,
                        HeadType = voucher.HeadType,
                        TaxAmount = voucher.TaxAmount,
                        TaxRate = voucher.TaxRate,
                        Reason = voucher.Reason,
                        //In case it's not BPV, ChequeDetails will be null, this way it will pass null to chequeno but avoiding null ref. error.
                        ChequeNo = voucher.ChequeDetails?.ChequeNo,
                        BankAccNo = voucher.ChequeDetails?.BankAccNo,
                        ChqBookNo = voucher.ChequeDetails?.ChqBookNo,
                        DueDate = voucher.DueDate,
                        Online_Vchr=voucher.Online_Vchr,
                        Online_Vchr_Type=voucher.Online_Vchr_Type,
                        Advising= voucher.Advising,
                        VendID = voucher.VendID,
                        EmpID = voucher.EmpID,
                        PaymentToEmp = voucher.PaymentToEmp
                    },
                    commandType: CommandType.StoredProcedure,
                    transaction: transaction
                    );

                //First Save Cheque Details in Cheque Table.
                //Get the SNo from Cheque and use it in vouchers for the Bank Account Heaad Only
                // 2. CALL THE LINE ITEM SP (TVP)
                // This call uses the VoucherID obtained in Step 1
                long? returnSNo = null;
                if (voucher.ChequeDetails != null)
                {
                    returnSNo = await InsertChequeDetailsAsync(db, transaction, voucher.ChequeDetails);
                }
                await InsertVoucherLineItemsAsync(db, transaction, voucher.LineItems, strVchrNo);
                // 3. Commit the Dapper Transaction
                if (isInternalConn) transaction.Commit();
                return strVchrNo;
            }
            catch (Exception ex)
            {
                if (existingTrans == null) transaction.Rollback();
                //transaction.Rollback();
                _logger.LogError(
                    ex, // Pass the exception object itself
                    "Database Transaction Failed for Voucher: {VoucherNumber} by User {Username} from {MachineName}. Error: {ErrorMessage}",
                    voucher.VchrNo, // Value for the first placeholder
                    voucher.UserName,
                    voucher.MachineName,
                    ex.Message
                );
                throw;
            }
             
            finally
            {
                if (existingTrans == null) transaction.Dispose();
                if (isInternalConn) db.Dispose();
            }
        }
        public async Task<string> ExecuteVoucherSave_Old(VoucherViewModel voucher)
        {
            
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                db.Open();
                // Use a Dapper Transaction to wrap both SP calls to ensure all or nothing
                using (var transaction = db.BeginTransaction())
                {
                    try
                    {
                        // 1. CALL THE HEADER SP (This SP contains the locking logic internally)
                        // It returns the new VoucherID
                        /*string strVchrNo = await db.QuerySingleAsync<string>(
                            "dbo.InsertVoucherHead_SP \r\n",
                             voucher, // Pass the whole model/object as parameters
                             commandType: CommandType.StoredProcedure,
                             transaction: transaction
                        );*/
                        string strVchrNo = await db.QuerySingleAsync<string>(
                            "dbo.InsertVoucherHead_SP",
                            new
                            {
                                VchrNo = voucher.VchrNo,
                                UserName = voucher.UserName,
                                MachineName = voucher.MachineName,
                                DT = voucher.DT,
                                Notes = voucher.Notes,
                                Payee = voucher.Payee,
                                TaxType = voucher.TaxType,
                                HeadType = voucher.HeadType,
                                TaxAmount = voucher.TaxAmount,
                                TaxRate = voucher.TaxRate,
                                Reason = voucher.Reason,
                                //In case it's not BPV, ChequeDetails will be null, this way it will pass null to chequeno but avoiding null ref. error.
                                ChequeNo = voucher.ChequeDetails?.ChequeNo,
                                BankAccNo = voucher.ChequeDetails?.BankAccNo,
                                ChqBookNo = voucher.ChequeDetails?.ChqBookNo,
                                DueDate=voucher.DueDate
                            },
                            commandType: CommandType.StoredProcedure,
                            transaction: transaction
                            );

                        //First Save Cheque Details in Cheque Table.
                        //Get the SNo from Cheque and use it in vouchers for the Bank Account Heaad Only
                        // 2. CALL THE LINE ITEM SP (TVP)
                        // This call uses the VoucherID obtained in Step 1
                        long? returnSNo = null;
                        if (voucher.ChequeDetails != null)
                        {
                            returnSNo = await InsertChequeDetailsAsync(db, transaction, voucher.ChequeDetails);
                        }


                        await InsertVoucherLineItemsAsync(db, transaction, voucher.LineItems, strVchrNo);

                        // 3. Commit the Dapper Transaction
                        transaction.Commit();
                        return strVchrNo;
                    }
                    catch (Exception ex)
                    {
                        {
                            transaction.Rollback();
                            _logger.LogError(
                                ex, // Pass the exception object itself
                                "Database Transaction Failed for Voucher: {VoucherNumber} by User {Username} from {MachineName}. Error: {ErrorMessage}",
                                voucher.VchrNo, // Value for the first placeholder
                                voucher.UserName,
                                voucher.MachineName,
                                ex.Message
                            );
                            throw;
                        }
                    }
                }
            }
        }
    private async Task InsertVoucherLineItemsAsync(IDbConnection db,IDbTransaction transaction,List<VoucherLineItemViewModel> lineItems,string vchrNo) // Note: Accepts the string VchrNo as the FK
        {
            // 1. Prepare the data for the TVP
            // We project the C# list into an anonymous type that matches the SQL Table Type structure.
            if (!lineItems.Any())
            {
                return;
            }

            DataTable dt = new DataTable();

            // Define columns based on the SQL Table Type. Names MUST match the SQL column names.
            dt.Columns.Add("SNo", typeof(double));    
            dt.Columns.Add("VDate", typeof(DateTime));    
            dt.Columns.Add("VchrNo", typeof(string));
            dt.Columns.Add("AccNo", typeof(string));
            dt.Columns.Add("Description", typeof(string));
            dt.Columns.Add("Debit", typeof(decimal));
            dt.Columns.Add("Credit", typeof(decimal));
            dt.Columns.Add("Balance", typeof(decimal));
            dt.Columns.Add("DpstSlip", typeof(string));
            dt.Columns.Add("CSNo", typeof(long));        
            dt.Columns.Add("CS_RefID", typeof(int));

            
            foreach (var line in lineItems)
            {
                if (line != null && (line.Debit > 0 || line.Credit > 0))
                {
                    dt.Rows.Add(
                        line.SNo,   //Need to update SNos
                        line.VDate,
                        vchrNo,    
                        line.AccNo,
                        line.Description,
                        line.Debit,
                        line.Credit,
                        line.Balance,
                        line.DpstSlip,
                        line.CSNo,
                        line.CS_RefID                        
                    );
                }
            }

            string spName = "dbo.InsertVoucherLineItems_SP";
            string tvpTypeName = "dbo.Vouchers_TT";

            // 3. Execute the Stored Procedure with the TVP
            await db.ExecuteAsync(
                spName,
                new
                {
                    @Vouchers_TT = dt.AsTableValuedParameter(tvpTypeName)
                },
                commandType: CommandType.StoredProcedure,
                transaction: transaction // Remember to update Balance and SNo in SP, may be I should do it in C# before starting transaction.
            );
        }
    private async Task<long?> InsertChequeDetailsAsync(IDbConnection db,IDbTransaction transaction,ChequeViewModel chequeDetails)
        {
            // If no cheque details are provided (e.g., cash voucher), return null
            if (chequeDetails == null)
            {
                return null;
            }

            // Use Dapper's QuerySingleAsync to execute the SP and retrieve the single integer ID.
            // Dapper maps the ChequeViewModel properties to the SP parameters.
            long CSNo = await db.QuerySingleAsync<long>(
                "dbo.InsertChequeDetails_SP",
                chequeDetails,
                commandType: CommandType.StoredProcedure,
                transaction: transaction // Ensures it runs within the main voucher transaction
            );

            return CSNo;
        }

        public async Task<List<VAD_Voucher_DetailViewModel>> GetPendingJVsByAccount(string accNo)
        {
            // Replicating: SELECT * FROM VLedger_Balance WHERE AccNo=... AND LEFT(VchrNo,2)='JV'...
            IDbConnection db = new SqlConnection(_connectionString);
            string sql = @"SELECT AccNo,AccTitle,SNo,VchrNo,Debit AS VchrAmount,AdjAmount,RemainBalance AS BalanceAmount,CS_RefID FROM VLedger_Balance " +
                "WHERE AccNo = @AccNo AND VchrNo LIKE 'JV%' AND Debit > 0 AND ISNULL(RemainBalance, 0) > 0 ORDER BY VDate DESC";
            
            var results = await db.QueryAsync<VAD_Voucher_DetailViewModel>(sql, new { AccNo = accNo });
            return results.ToList();
        }
        public async Task<Voucher_MainSaveResult> ExecuteVoucher_MainSave(Vouchers_MainViewModel voucher)
        {

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                db.Open();
                // Use a Dapper Transaction to wrap both SP calls to ensure all or nothing
                using (var transaction = db.BeginTransaction())
                {
                    try
                    {                        
                        Voucher_MainSaveResult headResult = await db.QueryFirstAsync<Voucher_MainSaveResult>(
                            "dbo.InsertVoucherApprovalHead_SP",
                            new
                            {
                                VchrNo = voucher.VchrNo,
                                UserName = voucher.UserName,
                                MachineName = voucher.MachineName,
                                DT = voucher.DT,
                                Remarks = voucher.Remarks,
                                DueDate=voucher.DueDate,
                                SlipNo=voucher.SlipNo,
                                ReceivedFrom=voucher.ReceivedFrom,
                                Attachment_Type=voucher.Attachment_Type,
                                PDFAttachments=voucher.PDFAttachments,
                                ForwardBy=voucher.UserName,
                                ForwardTo=voucher.ForwardTo,
                                TaxType = voucher.TaxType,
                                HeadType = voucher.HeadType,
                                TaxAmount = voucher.TaxAmount,
                                TaxRate = voucher.TaxRate,
                                Reason = voucher.Reason
                                
                            },
                            commandType: CommandType.StoredProcedure,
                            transaction: transaction
                            );
                        
                        int newEntryID=headResult.EntryID;
                        string strVchrNo=headResult.VchrNo;

                        string sqlLine = @"
                            INSERT INTO Vouchers_Approval_Detail (RefID,VchrNo,Accno,Description,Debit,Credit,Balance,CS_RefID)
                            VALUES (@RefID,@VchrNo,@AccNo,@Description,@Debit,@Credit,@Balance,@CS_RefID);
                            SELECT CAST(SCOPE_IDENTITY() as int);";

                        foreach (var line in voucher.LineItems) 
                        {
                            int VAD_EntryID = await db.ExecuteScalarAsync<int>(sqlLine, new
                            {
                                RefID = newEntryID,
                                VchrNo=strVchrNo,
                                line.AccNo,
                                line.Description,
                                line.Debit,
                                line.Credit,
                                line.Balance,
                                line.CS_RefID
                            }, transaction);
                            
                            if (line.AdjustmentDetails != null && line.AdjustmentDetails.Any())
                            {
                                string sqlAdj = @"
                                        INSERT INTO VAD_Voucher_Detail(RefID,VchrNo,SNo,AccNo,Amount,BalanceAmount,AdjAmount,RemainAmount,VchrAmount)
                                        VALUES (@RefID,@VchrNo,@SNo,@AccNo,@Amount,@BalanceAmount,@AdjAmount,@RemainAmount,@VchrAmount);";

                                foreach (var adj in line.AdjustmentDetails)
                                {
                                    await db.ExecuteAsync(sqlAdj, new
                                    {
                                        RefID = VAD_EntryID,                                        
                                        adj.VchrNo,
                                        adj.Sno,
                                        adj.AccNo,
                                        adj.Amount, // This is the user entered adjustment
                                        adj.BalanceAmount,
                                        adj.AdjAmount,
                                        adj.RemainAmount,
                                        adj.VchrAmount                                        
                                    }, transaction);
                                }
                            }
                        }
                        
                        // 3. Commit the Dapper Transaction
                        transaction.Commit();
                        return new Voucher_MainSaveResult
                        {
                            VchrNo = strVchrNo,
                            EntryID = newEntryID
                        };
                    }
                    catch (Exception ex)
                    {
                        {
                            transaction.Rollback();
                            _logger.LogError(
                                ex, // Pass the exception object itself
                                "Database Transaction Failed for Voucher: {VoucherNumber} by User {Username} from {MachineName}. Error: {ErrorMessage}",
                                voucher.VchrNo, // Value for the first placeholder
                                voucher.UserName,
                                voucher.MachineName,
                                ex.Message
                            );
                            throw;
                        }
                    }
                }
            }
        }
        public async Task<int> ExecuteOnlinePaymentVoucherSave(Online_PVModelView header, List<Online_PV_DetailsModelView> details)
        {
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();
            using var transaction = connection.BeginTransaction();

            try
            {
                // 1. Insert Header and get the Identity
                string headerSql = @"INSERT INTO Online_PV (DT,SIAL_Bank_Acc_No,Debit_Acc_No,Tax_Acc_No,UserName,MachineName,Transaction_Type,Remarks,BOD_Meetings_RefID) 
                             VALUES (@DT,@SIAL_Bank_Acc_No,@Debit_Acc_No,@Tax_Acc_No,@UserName,@MachineName,@Transaction_Type,@Remarks,@BOD_Meetings_RefID);
                             SELECT SCOPE_IDENTITY();";

                int pvID = await connection.ExecuteScalarAsync<int>(headerSql, header, transaction);

                // 2. Insert Details
                string detailSql = @"INSERT INTO Online_PV_Detail (OL_PV_RefID,SD_RefID,Bank_Acc_No,Bank_Code,Bank_Acc_Title,Bank_Name,Debit,Tax)
                             VALUES (@OL_PV_RefID, @SD_RefID, @Bank_Acc_No, @Bank_Code, @Bank_Acc_Title, @Bank_Name, @Debit, @Tax)";

                foreach (var item in details)
                {
                    item.OL_PV_RefID = pvID; // Link to the header we just created
                    await connection.ExecuteAsync(detailSql, item, transaction);
                }

                transaction.Commit();
                return pvID;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                _logger.LogError(ex, "Failed to save Online PV staging data.");
                throw;
            }
        }
        public async Task<string> ExecuteBODMeetingPaymentVoucherSave(VoucherViewModel voucher,BOD_Meeting_Payment_List_ViewModel BODPLVM,double chqBookNo, double SNo_DebitLine)
        {
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();
            using var transaction = connection.BeginTransaction();

            try
            {
                string strVchrNo=await ExecuteVoucherSave(voucher, connection, transaction);
                // 1. Insert Header and get the Identity
                string strUpdate = @"UPDATE Online_PV SET VchrNo=@VchrNo WHERE EntryID=@EntryID";
                await connection.ExecuteScalarAsync(strUpdate, new { VchrNo = strVchrNo,EntryID= BODPLVM.EntryID }, transaction);
                strUpdate = "UPDATE VoucherInfo SET Notes=@Notes,Online_Vchr=1,Online_Vchr_Type=1 WHERE VchrNo=@VchrNo";
                await connection.ExecuteScalarAsync(strUpdate, new { Notes = BODPLVM.Remarks, VchrNo = strVchrNo }, transaction);

                if (BODPLVM.Transaction_Type == 1)
                {
                    double dVoucher_SNo = SNo_DebitLine;//await _db.GetSingleValueAsync<double>("SNo", "Vouchers", " WHERE VchrNo='" + strVchrNo + "' AND AccNo='" + BODPLVM.Debit_Acc_No + "'", connection, transaction);

                    var details = await connection.QueryAsync<dynamic>(
                    "SELECT * FROM Online_PV_Detail WHERE OL_PV_RefID = @ID",
                    new { ID = BODPLVM.EntryID }, transaction);

                    string strChqType = "Open"; //Payees A/C Only
                    foreach (var item in details)
                    {
                        // Replicate: GetNextChqNo & getNextSno logic
                        string strChqNo = await GetNextChqNo(BODPLVM.SIAL_Bank_Acc_No, chqBookNo,1,1,connection, transaction);
                        long SNo = await GetNextSNo(voucher.DT, "Cheque",connection, transaction);
                        
                        string strChqDesc = "Chq#"+ strChqNo+" BOD Meeting Fee Dtd."+voucher.DT.ToString("dd.MM.yy");
                        // Update Detail Staging
                        await connection.ExecuteAsync(
                            "UPDATE Online_PV_Detail SET ChqNo=@Chq,Advising=@Advising WHERE EntryID=@DetID",
                            new { Chq = strChqNo,Advising=strChqDesc, DetID = item.EntryID }, transaction);

                        // Insert into main Cheque table
                        await connection.ExecuteAsync(
                            "INSERT INTO Cheque (SNo,CDate,AccNo,Description,Amount,BankID,ChequeNo,ChqBookNo,ChequeType,ChequeDate,Posted,Payment,Bounced,Payee,BankAccNo,Vouchers_SNo) " +
                            "VALUES (@SNo,@DT,@Acc,@Description,@Amt,@BankID,@ChqNo,@ChqBookNo,@ChequeType,@DT,0,1,0,@Payee,@BankAccNo,@VchrSNo)",
                            new
                            {
                                SNo = SNo,
                                DT=voucher.DT,
                                Acc = BODPLVM.Debit_Acc_No,
                                Description= strChqDesc,
                                Amt = item.Debit - item.Tax,
                                BankID=75,
                                ChqNo = strChqNo,
                                ChqBookNo = chqBookNo,
                                ChequeType= strChqType,
                                Payee = item.DirectorName,
                                BankAccNo=BODPLVM.SIAL_Bank_Acc_No,
                                VchrSNo= dVoucher_SNo

                            }, transaction);

                        // Mark Chq as Issued
                        await connection.ExecuteAsync("UPDATE ChqList SET Issued=1 WHERE ChqBookNo=@ChqBookNo AND ChqNo=@Chq", new { ChqBookNo=chqBookNo, Chq = strChqNo }, transaction);
                    }
                }                

                transaction.Commit();
                return strVchrNo;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                _logger.LogError(ex, "Failed to save Online PV staging data.");
                throw;
            }
        }

    }
}


