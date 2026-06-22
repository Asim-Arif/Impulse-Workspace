using System.Text;
using System.Xml.Linq;

namespace CIP.Services.Integrations
{
    public class MeezanBankService : IMeezanBankService
    {
        private readonly HttpClient _httpClient;
        private readonly IConfiguration _config;
        private readonly IDBHelperService _db;
        private IAuditService _auditService { get; set; }
        private readonly IAccountReportingAccess _accountReportingAccess;
        private string strCurrentUserName;
        private string strMachineName;
        public MeezanBankService(HttpClient httpClient, IConfiguration config, IDBHelperService db, IAccountReportingAccess accountReportingAccess, IAuditService auditService)
        {
            _httpClient = httpClient;
            _config = config;
            _db = db;
            _accountReportingAccess = accountReportingAccess;
            _auditService= auditService;

            strCurrentUserName = "later";//_auditService.GetCurrentUserName();
            strMachineName=_auditService.GetClientIpAddress();
        }

        public async Task<string> FetchTitleMeezan(string strDebitAccount, string strCreditAccount, string strTransactionType, string strBankCode = "")
        {
            // 1. Fetch Configuration (Equivalent to getGeneralDataValue)
            string strServer = await _db.getGeneralDataValueAsync("Meezan_QP_Server_Address");
            string strUserID = await _db.getGeneralDataValueAsync("Meezan_QP_UserID");
            string strRequestAtt = await _db.getGeneralDataValueAsync("Meezan_QP_Request_Attribute");

            string headerSetting = strTransactionType == "InterBankFundTransfer_TitleFetch"
                ? "Meezan_QP_Request_Head_IBFT_Title_Fetch"
                : "Meezan_QP_Request_Head_IFT_Title_Fetch";

            string strHeader = await _db.getGeneralDataValueAsync(headerSetting);

            string strPassword = await _db.getDatabasePasswordAsync("Meezan_Password");
            string strAuthCode = await _db.getDatabasePasswordAsync("Meezan_AuthenticationCode");
            
            // 3. Construct SOAP Envelope using XDocument (Modern XML handling)
            XNamespace soap = "http://schemas.xmlsoap.org/soap/envelope/";
            XNamespace xsi = "http://www.w3.org/2001/XMLSchema-instance";
            XNamespace xsd = "http://www.w3.org/2001/XMLSchema";
            XNamespace cas = strRequestAtt;

            var soapRequest = new XDocument(
                new XDeclaration("1.0", "utf-8", "yes"),
                new XElement(soap + "Envelope",
                    new XAttribute(XNamespace.Xmlns + "soap", soap),
                    new XAttribute(XNamespace.Xmlns + "xsi", xsi),
                    new XAttribute(XNamespace.Xmlns + "xsd", xsd),
                    new XAttribute(XNamespace.Xmlns + "cas", cas),
                    new XElement(soap + "Body",
                        new XElement(cas + strTransactionType,
                            new XElement(cas + "TransactionDetail",
                                new XElement(cas + "UserID", strUserID),
                                new XElement(cas + "Password", strPassword),
                                new XElement(cas + "AuthenticationCode", strAuthCode),
                                new XElement(cas + "DebitAccountNumber", strDebitAccount),
                                new XElement(cas + "CreditAccountNumber", strCreditAccount),
                                strTransactionType == "InterBankFundTransfer_TitleFetch" ? new XElement(cas + "BankCode", strBankCode) : null
                            )
                        )
                    )
                )
            );

            // 4. Send Request
            var request = new HttpRequestMessage(HttpMethod.Post, strServer);
            request.Content = new StringContent(soapRequest.ToString(), Encoding.UTF8, "text/xml");
            request.Headers.Add("SOAPAction", strHeader);

            var response = await _httpClient.SendAsync(request);
            string responseText = await response.Content.ReadAsStringAsync();

            // 5. Parse Response (You'll likely want to extract a specific node here)
            return responseText;
        }
        public async Task<bool> PostToMeezanPortal(int entryID, int transactionType)
        {
            // 1. Fetch Credentials & Settings (Replaces getGeneralDataValue)
            var serverUrl = await _db.getGeneralDataValueAsync("Meezan_QP_Server_Address");
            var userID = await _db.getGeneralDataValueAsync("Meezan_QP_UserID");
            
            string strPassword = await _db.getDatabasePasswordAsync("Meezan_Password");
            string strAuthCode = await _db.getDatabasePasswordAsync("Meezan_AuthenticationCode");
            
            // 2. Get the Debit Account 
            // Assuming you have the Master record to get the Bank Account Tag            
            string debitAccNo = await _db.GetSingleValueAsync<string>("BankAccNo", "VOnline_PV", $" WHERE EntryID={entryID}");
            // 3. Fetch Details
            var details = await _accountReportingAccess.GetBODMeetingPaymentDetail(entryID);

            foreach (var item in details)
            {
                // Replicate VB6 Logic for Transaction Types
                string transactionTypeName = "";
                string header = "";
                string beneficiaryName = item.DirectorName; // Default
                string remarks = item.FolioNo.ToString();
                
                string strCustomerReference = "DIR" + item.EntryID.ToString();
                switch (transactionType)
                {
                    case 0: // Online Funds Transfer
                        bool isMeezan = item.Bank_Name.Contains("meezan", StringComparison.OrdinalIgnoreCase);
                        transactionTypeName = isMeezan ? "InternalFundsTransfe" : "InterBankFundTransfer";
                        header = await _db.getGeneralDataValueAsync(isMeezan ? "Meezan_QP_Request_Header" : "Meezan_QP_Request_Header_IBFT");
                        // Note: Fetch_Title_Meezan_Fast should be an async helper
                        beneficiaryName = await FetchTitleMeezan(debitAccNo, item.Bank_Acc_No, transactionTypeName, item.Bank_Code);
                        
                        break;

                    case 1: // Corporate Cheque
                        transactionTypeName = "CorporateCheque";
                        header = await _db.getGeneralDataValueAsync("Meezan_QP_Request_Header_Corporate_Cheque");
                        remarks = item.Advising;
                        break;                        
                }

                // 4. Call the Bank API (The modern version of Post_Meezan_QP)
                int netAmount = item.Debit - item.Tax;
                var myRequest = new MeezanRequestModel
                {
                    ServerUrl = serverUrl,
                    UserID = userID,
                    Password=strPassword,
                    AuthCode=strAuthCode,
                    CustomerReference= strCustomerReference,
                    DebitAccount=debitAccNo,
                    Amount = netAmount,
                    BeneficiaryName = beneficiaryName,
                    CreditAccount=item.Bank_Acc_No,
                    IType= transactionType,
                    TransactionType=transactionTypeName,
                    BankCode=item.Bank_Code,
                    Remarks=remarks,                    
                };

                string responseXml = await PostMeezanQPAsync(myRequest);

                //string responseXml = await _bankApi.SendMeezanRequest(serverUrl, header, userID, strPassword, strAuthCode, debitAccNo, item, transactionTypeName, netAmount);

                // 5. Parse XML Response (Using XDocument instead of string manipulation)
                var xmlDoc = XDocument.Parse(responseXml);
                string resCode = xmlDoc.Root.Element("ResponseCode")?.Value ?? "ERR";
                string resDesc = xmlDoc.Root.Element("ResponseDescription")?.Value ?? "No Response";
                string resTxn = xmlDoc.Root.Element("UnqiueTxnNo")?.Value ?? "";

                // 6. Update Detail Staging Table
                await _db.ExecuteAsync(@"UPDATE Online_PV_Detail SET UserName_Post=@User, MachineName_Post=@Machine, DTEntry_Post=GETDATE(), Status=1, 
                                    Response_Raw=@Raw, Response_Code=@Code, Response_Description=@Desc, Response_UnqiueTxnNo=@Txn 
                                    WHERE EntryID=@DetID",
                    new
                    {
                        User = strCurrentUserName,
                        Machine = strMachineName,
                        Raw = responseXml,
                        Code = resCode,
                        Desc = resDesc,
                        Txn = resTxn,
                        DetID = item.EntryID                        
                    });
                
            }

            // 7. Final Master Update
            await _db.ExecuteAsync(@"UPDATE Online_PV SET Status=1, UserName_Approved=@User,MachineName_Approved=@Machine,DTEntry_Approved=GETDATE() 
                                    WHERE EntryID=@ID",
                new { User = strCurrentUserName, Machine = strMachineName, ID = entryID });

            return true;
        }
        

    public async Task<string> PostMeezanQPAsync(MeezanRequestModel req)
    {
        // 1. Setup the HTTP Client with SSL bypass (Option 2, 13056 equivalent)
        var handler = new HttpClientHandler
        {
            ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true
        };

        using var client = new HttpClient(handler);
        client.DefaultRequestHeaders.Add("SOAPAction", req.Header);

        // 2. Build the SOAP XML Envelope using LINQ to XML
        XNamespace soap = "http://schemas.xmlsoap.org/soap/envelope/";
        XNamespace xsi = "http://www.w3.org/2001/XMLSchema-instance";
        XNamespace xsd = "http://www.w3.org/2001/XMLSchema";
        XNamespace ns = req.RequestAttribute; // The xmlns for the call

        var soapEnvelope = new XElement(soap + "Envelope",
            new XAttribute(XNamespace.Xmlns + "soap", soap.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "xsi", xsi.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "xsd", xsd.NamespaceName),
            new XElement(soap + "Body",
                new XElement(ns + req.TransactionType,
                    new XAttribute("xmlns", ns.NamespaceName),
                    new XElement(ns + "TransactionDetail",
                        new XAttribute("xmlns", ns.NamespaceName),
                        new XElement(ns + "UserID", req.UserID),
                        new XElement(ns + "Password", req.Password),
                        new XElement(ns + "AuthenticationCode", req.AuthCode),
                        new XElement(ns + "CustomerReferenceNumber", req.CustomerReference),
                        new XElement(ns + "DebitAccountNumber", req.DebitAccount),

                        // Logic for Credit Account (iType 0 or 4)
                        (req.IType == 0 || req.IType == 4) ? new XElement(ns + "CreditAccountNumber", req.CreditAccount) : null,

                        new XElement(ns + "TransactionAmount", req.Amount),

                        // Logic for Bank Code
                        (req.TransactionType == "InterBankFundTransfer" || req.TransactionType == "RealTimeGrossSettlement")
                            ? new XElement(ns + "BankCode", req.BankCode) : null,

                        // Logic for Corporate Cheques (iType 1, 2, 3)
                        (req.IType >= 1 && req.IType <= 3) ? new[] {
                        new XElement(ns + "DeliverTo", "1"),
                        new XElement(ns + "PrintLocation", "skt")
                        } : null,

                        // Logic for E-Biz Plus (iType 2)
                        (req.IType == 2) ? new[] {
                        new XElement(ns + "BeneAddress", req.BeneficiaryAddress),
                        new XElement(ns + "BeneContact", req.BeneficiaryContactNo),
                        new XElement(ns + "CNIC", req.IdentificationNumber),
                        new XElement(ns + "Purpose", req.Purpose)
                        } : null,

                        new XElement(ns + "Narration", req.Remarks),

                        // Logic for BeneName (All types except potentially specific exclusions)
                        new XElement(ns + "BeneName", req.BeneficiaryName)
                    )
                )
            )
        );

        string xmlString = soapEnvelope.ToString();

        // 3. Save Request to DB for logging (as you did in VB6)
        await _db.ExecuteAsync("UPDATE Online_PV_Detail SET Request_Raw=@Raw WHERE EntryID=@ID",
                                new { Raw = xmlString, ID = req.DetailEntryID });

        // 4. Send the Request
        var content = new StringContent(xmlString, Encoding.UTF8, "text/xml");

        try
        {
            var response = await client.PostAsync(req.ServerUrl, content);
            return await response.Content.ReadAsStringAsync();
        }
        catch (Exception ex)
        {
            return $"<Error>{ex.Message}</Error>";
        }
    }        
}
    public class MeezanRequestModel
    {
        // Connection & Auth
        public string ServerUrl { get; set; } = string.Empty;
        public string Header { get; set; } = string.Empty;
        public string RequestAttribute { get; set; } = string.Empty;
        public string UserID { get; set; } = string.Empty;
        public string Password { get; set; } = string.Empty;
        public string AuthCode { get; set; } = string.Empty;

        // Transaction Identification
        public string TransactionType { get; set; } = string.Empty; // e.g., "InternalFundsTransfe"
        public int IType { get; set; } // iTransaction_Type (0 to 4)
        public string CustomerReference { get; set; } = string.Empty;
        public long DetailEntryID { get; set; }

        // Financial Data
        public string DebitAccount { get; set; } = string.Empty;
        public string CreditAccount { get; set; } = string.Empty;
        public int Amount { get; set; }
        public string Remarks { get; set; } = string.Empty; // Narration

        // Beneficiary Info
        public string BeneficiaryName { get; set; } = string.Empty;
        public string BeneficiaryAddress { get; set; } = string.Empty;
        public string BeneficiaryContactNo { get; set; } = string.Empty;
        public string IdentificationNumber { get; set; } = string.Empty; // CNIC
        public string Purpose { get; set; } = string.Empty;

        // Bank Routing
        public string BankCode { get; set; } = string.Empty;
        public string BranchCode { get; set; } = string.Empty;
        public string BankName { get; set; } = string.Empty;
    }
}

