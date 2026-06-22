using ClosedXML.Excel;
using ClosedXML.Excel;
using DataAccessLibrary.DAC.Accounts;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml.Office;
using Renci.SshNet;
using System.Data;
using System.IO;

namespace CIP.Services.Integrations
{
    public class HabibBankServices : IHabibBankServices
    {
        private readonly ILogger<FaysalBankServices> _logger;
        private readonly HttpClient _httpClient;
        private readonly IConfiguration _config;
        private readonly IDBHelperService _db;
        private IAuditService _auditService { get; set; }
        private readonly IAccountReportingAccess _accountReportingAccess;
        private string strCurrentUserName;
        private string strMachineName;
        
        public HabibBankServices(HttpClient httpClient, IConfiguration config, IDBHelperService db, IAccountReportingAccess accountReportingAccess, IAuditService auditService, ILogger<FaysalBankServices> logger)
        {
            _httpClient = httpClient;
            _config = config;
            _db = db;
            _accountReportingAccess = accountReportingAccess;
            _auditService = auditService;
            _logger = logger;
            strCurrentUserName = _auditService.GetCurrentUserName();
            strMachineName = _auditService.GetClientIpAddress();
        }
        
        public async Task<int> ExportExcelAsync(int entryId)
        {
            var details = await _accountReportingAccess.GetBODMeetingPaymentDetail(entryId);

            int iCount_FT= details.Count(x => x.Bank_Code_Name != null && x.Bank_Code_Name == "FBL" && x.Transaction_Type==0);
            int iCount_IBFT = details.Count(x => x.Bank_Code_Name != null && x.Bank_Code_Name != "FBL" && x.Transaction_Type == 0);
            int iCount_CC = details.Count(x => x.Transaction_Type == 1);

            bool isUploaded = false;
            string fileName;
            int row = 2;
            if (iCount_FT > 0)
            {
                row = 2;
                using var workbook_FT = new XLWorkbook();
                var worksheet_FT = workbook_FT.Worksheets.Add("Sheet1");
                worksheet_FT.Cell(1, 1).Value = "Beneficiary First Name";
                worksheet_FT.Cell(1, 2).Value = "Beneficiary Account No";
                worksheet_FT.Cell(1, 3).Value = "Transaction Amount";
                worksheet_FT.Cell(1, 4).Value = "Reference # 1";
                worksheet_FT.Cell(1, 5).Value = "Beneficiary Account Title";
                foreach (var item in details.Where(x => x.Bank_Code_Name != null && x.Bank_Code_Name == "FBL" && x.Transaction_Type == 0)) //Funds Transfer
                {
                    worksheet_FT.Cell(row, 1).Value = item.Bank_Acc_Title;
                    worksheet_FT.Cell(row, 2).Value = item.Bank_Acc_No;
                    worksheet_FT.Cell(row, 3).Value = item.Debit - item.Tax;
                    worksheet_FT.Cell(row, 4).Value = item.EntryID;
                    worksheet_FT.Cell(row, 5).Value = item.Bank_Acc_Title;
                    row++;
                }

                worksheet_FT.Columns().AdjustToContents();
                // 3. Return as Byte Array for SFTP upload
                using var stream_FT = new MemoryStream();
                workbook_FT.SaveAs(stream_FT);
                byte[] fileBytes_FT = stream_FT.ToArray();
                
                fileName = "SIALAIFT_Faysal_Bank_FT_" + entryId.ToString() + ".xlsx";
                isUploaded = await UploadToBankAsync(fileName, fileBytes_FT);

                if (isUploaded)
                {
                    await _db.ExecuteAsync("UPDATE Online_PV_Detail SET UserName_Post=@UserName,MachineName_Post=@MachineName,DTEntry_Post=getDATE(),Status=1, Response_Raw=@FileName " +
                                            "WHERE EntryID IN(SELECT EntryID FROM VOnline_PV_Detail WHERE OL_PV_RefID=@ID AND Bank_Code_Name='FBL' AND Transaction_Type=0)",
                            new { UserName=strCurrentUserName,MachineName=strMachineName, FileName = fileName, ID = entryId });
                }
            }

            if (iCount_IBFT > 0)
            {
                row = 2;
                using var workbook_IBFT = new XLWorkbook();
                var worksheet_IBFT = workbook_IBFT.Worksheets.Add("Sheet1");
                worksheet_IBFT.Cell(1, 1).Value = "Beneficiary First Name";
                worksheet_IBFT.Cell(1, 2).Value = "Beneficiary Account No";
                worksheet_IBFT.Cell(1, 3).Value = "Bank";
                worksheet_IBFT.Cell(1, 4).Value = "Transaction Amount";
                worksheet_IBFT.Cell(1, 5).Value = "Reference # 1";
                worksheet_IBFT.Cell(1, 6).Value = "Reference # 9";
                worksheet_IBFT.Cell(1, 7).Value = "Beneficiary Account Title";
                foreach (var item in details.Where(x => x.Bank_Code_Name != null && x.Bank_Code_Name != "FBL" && x.Transaction_Type == 0))  //Interbank Funds Transfer
                {
                    worksheet_IBFT.Cell(row, 1).Value = item.Bank_Acc_Title;
                    worksheet_IBFT.Cell(row, 2).Value = item.Bank_Acc_No;
                    worksheet_IBFT.Cell(row, 3).Value = item.Bank_Code_Name;
                    worksheet_IBFT.Cell(row, 4).Value = item.Debit - item.Tax;
                    worksheet_IBFT.Cell(row, 5).Value = item.EntryID;
                    worksheet_IBFT.Cell(row, 6).Value = item.EntryID;
                    worksheet_IBFT.Cell(row, 7).Value = item.Bank_Acc_Title;
                    row++;
                }
                worksheet_IBFT.Columns().AdjustToContents();
                // 3. Return as Byte Array for SFTP upload
                using var stream_IBFT = new MemoryStream();
                workbook_IBFT.SaveAs(stream_IBFT);
                byte[] fileBytes_IBFT = stream_IBFT.ToArray();

                fileName = "SIALIBFT_Faysal_Bank_IBFT_" + entryId.ToString() + ".xlsx";
                isUploaded = await UploadToBankAsync(fileName, fileBytes_IBFT);

                if (isUploaded)
                {
                    await _db.ExecuteAsync("UPDATE Online_PV_Detail SET UserName_Post=@UserName,MachineName_Post=@MachineName,DTEntry_Post=getDATE(),Status=1, Response_Raw=@FileName " +
                                            "WHERE EntryID IN(SELECT EntryID FROM VOnline_PV_Detail WHERE OL_PV_RefID=@ID AND Bank_Code_Name<>'FBL' AND Transaction_Type=0)",
                            new { UserName = strCurrentUserName, MachineName = strMachineName, FileName = fileName, ID = entryId });
                }
            }

            if (iCount_CC > 0)
            {
                row = 2;
                using var workbook_CC = new XLWorkbook();
                var worksheet_CC = workbook_CC.Worksheets.Add("Sheet1");
                worksheet_CC.Cell(1, 1).Value = "Beneficiary First Name";
                worksheet_CC.Cell(1, 2).Value = "Beneficiary Address";
                worksheet_CC.Cell(1, 3).Value = "Transaction Amount";
                worksheet_CC.Cell(1, 4).Value = "Print Location";
                worksheet_CC.Cell(1, 5).Value = "Instrument No";
                worksheet_CC.Cell(1, 6).Value = "Reference # 1";
                worksheet_CC.Cell(1, 7).Value = "Reference # 2";
                worksheet_CC.Cell(1, 8).Value = "Reference # 3";
                foreach (var item in details.Where(x => x.Transaction_Type == 1))   //Corporate Cheque
                {
                    string vchrNo = item.VchrNo?.ToString().Replace("-", "") ?? "";
                    worksheet_CC.Cell(row, 1).Value = item.Bank_Acc_Title;
                    worksheet_CC.Cell(row, 2).Value = "Sialkot";
                    worksheet_CC.Cell(row, 3).Value = item.Debit-item.Tax;
                    worksheet_CC.Cell(row, 4).Value = "SIALK";
                    worksheet_CC.Cell(row, 5).Value = item.ChqNo;
                    worksheet_CC.Cell(row, 6).Value = item.EntryID;
                    worksheet_CC.Cell(row, 7).Value = vchrNo;
                    worksheet_CC.Cell(row, 8).Value = item.FolioNo;
                    row++;
                }
                worksheet_CC.Columns().AdjustToContents();
                // 3. Return as Byte Array for SFTP upload
                using var stream_CC = new MemoryStream();
                workbook_CC.SaveAs(stream_CC);
                byte[] fileBytes_CC = stream_CC.ToArray();

                fileName = "SIALICHQ_Faysal_Bank_CQ_" + entryId.ToString() + ".xlsx";
                string localPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);
                await File.WriteAllBytesAsync(localPath, fileBytes_CC);
                
                isUploaded = await UploadToBankAsync(fileName, fileBytes_CC);

                if (isUploaded)
                {
                    await _db.ExecuteAsync("UPDATE Online_PV_Detail SET UserName_Post=@UserName,MachineName_Post=@MachineName,DTEntry_Post=getDATE(),Status=1, Response_Raw=@FileName " +
                                            "WHERE EntryID IN(SELECT EntryID FROM VOnline_PV_Detail WHERE OL_PV_RefID=@ID AND Transaction_Type=1)",
                            new { UserName = strCurrentUserName, MachineName = strMachineName, FileName = fileName, ID = entryId });
                }
            }
            return iCount_FT + iCount_IBFT + iCount_CC;
        }
        public async Task<bool> UploadToBankAsync(string fileName, byte[] fileContent)
        {
            // These should be in your appsettings.json or GeneralData table
            
            string host =await _db.getGeneralDataValueAsync("Faysal_Bank_Host_Name");
            string username = await _db.getGeneralDataValueAsync("Faysal_Bank_User");
            int iPortNo = int.Parse(await _db.getGeneralDataValueAsync("Faysal_Bank_Port_No"));
            string password =await _db.getDatabasePasswordAsync("Faysal_Bank_Password");
            string remotePath = await _db.getGeneralDataValueAsync("Faysal_Bank_Folder_Path");
            string fullPath = $"{remotePath.TrimEnd('/')}/{fileName}";
            try
            {
                using (var client = new SftpClient(host,iPortNo, username, password))
                {
                    // Connect to Faysal Bank SFTP
                    await Task.Run(() => client.Connect());

                    using (var ms = new MemoryStream(fileContent))
                    {
                        // Upload the byte array as a file stream
                        await Task.Run(() => client.UploadFile(ms, fullPath));
                    }

                    client.Disconnect();
                    return true;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"SFTP Upload Failed for {fileName}");
                throw; // Re-throw to handle in the main loop
            }
        }
    }
}
