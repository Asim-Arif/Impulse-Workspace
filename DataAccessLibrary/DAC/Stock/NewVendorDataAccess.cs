using Dapper;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Stock
{
    public class NewVendorDataAccess : INewVendorDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<NewVendorDataAccess> _logger;

        public NewVendorDataAccess(IConfiguration configuration, ILogger<NewVendorDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection") 
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        public async Task<List<NewVendorViewModel>> GetVendorsAsync()
        {
            using IDbConnection connection = new SqlConnection(_connectionString);
            string sql = @"
                SELECT v.*, a.AccTitle as VendorName, a.SubAccOf as ParentAccNo, a.openbal as OpeningBalance, 
                       a.Balance as CurrentBalance,
                       CASE WHEN a.Active = 1 THEN 0 ELSE 1 END as Inactive 
                FROM Venders v
                LEFT JOIN Accounts a ON v.AccNo = a.AccNo
                ORDER BY v.VendID DESC";
            
            var result = await connection.QueryAsync<NewVendorViewModel>(sql);
            return result.ToList();
        }

        public async Task<NewVendorViewModel> GetVendorByIdAsync(int id)
        {
            using IDbConnection connection = new SqlConnection(_connectionString);
            string sql = @"
                SELECT v.*, a.AccTitle as VendorName, a.SubAccOf as ParentAccNo, a.openbal as OpeningBalance, 
                       CASE WHEN a.openbal >= 0 THEN 'DR' ELSE 'CR' END as OpeningBalanceType,
                       CASE WHEN a.Active = 1 THEN 0 ELSE 1 END as Inactive
                FROM Venders v
                LEFT JOIN Accounts a ON v.AccNo = a.AccNo
                WHERE v.VendID = @Id";
            
            var vendor = await connection.QueryFirstOrDefaultAsync<NewVendorViewModel>(sql, new { Id = id });
            
            if (vendor != null && vendor.OpeningBalance < 0)
            {
                vendor.OpeningBalance = Math.Abs(vendor.OpeningBalance);
            }
            
            return vendor ?? new NewVendorViewModel();
        }

        public async Task<string> GetNextMakerNoAsync()
        {
            using IDbConnection connection = new SqlConnection(_connectionString);
            string sql = "SELECT ISNULL(COUNT(VendID), 0) + 1 FROM Venders";
            int nextId = await connection.ExecuteScalarAsync<int>(sql);
            return $"SUP-{nextId:D4}";
        }

        public async Task<int> InsertVendorAsync(NewVendorViewModel model)
        {
            using IDbConnection connection = new SqlConnection(_connectionString);
            connection.Open();
            using IDbTransaction transaction = connection.BeginTransaction();

            try
            {
                string accType = !string.IsNullOrEmpty(model.ParentAccNo) && model.ParentAccNo.Length >= 2 
                    ? model.ParentAccNo.Substring(0, 2) 
                    : "24";

                // Generate next AccNo for the selected category using standard ERP logic
                string newAccNo = await connection.ExecuteScalarAsync<string>(
                    "SELECT dbo.GetNextAccno(@AccountName, @AccType, @ParentAccount, 0)",
                    new {
                        AccountName = model.VendorName ?? string.Empty,
                        AccType = accType,
                        ParentAccount = model.ParentAccNo
                    },
                    transaction);

                // Robust fallback in case function returns null or empty
                if (string.IsNullOrWhiteSpace(newAccNo))
                {
                    string prefix = "00";
                    if (!string.IsNullOrWhiteSpace(model.VendorName))
                    {
                        char firstChar = char.ToUpperInvariant(model.VendorName.Trim()[0]);
                        if (firstChar >= 'A' && firstChar <= 'Z')
                        {
                            prefix = (firstChar - 64).ToString("00");
                        }
                    }

                    // Query existing child account numbers under this parent safely without SQL casting
                    var existingAccNos = (await connection.QueryAsync<string>(
                        "SELECT AccNo FROM Accounts WHERE SubAccOf = @ParentAccNo",
                        new { ParentAccNo = model.ParentAccNo },
                        transaction)).ToList();

                    int maxSeq = 0;
                    foreach (var acc in existingAccNos)
                    {
                        if (string.IsNullOrEmpty(acc)) continue;
                        int lastDash = acc.LastIndexOf('-');
                        if (lastDash >= 0 && lastDash + 1 < acc.Length)
                        {
                            string suffix = acc.Substring(lastDash + 1);
                            // Suffix format is {Prefix 2 digits}{Seq 3 digits} e.g. 03009
                            if (suffix.Length == 5 && suffix.StartsWith(prefix))
                            {
                                if (int.TryParse(suffix.Substring(2), out int seq) && seq > maxSeq)
                                {
                                    maxSeq = seq;
                                }
                            }
                            else if (suffix.Length <= 3 && int.TryParse(suffix, out int simpleSeq) && simpleSeq > maxSeq)
                            {
                                maxSeq = simpleSeq;
                            }
                        }
                    }

                    newAccNo = $"{model.ParentAccNo}-{prefix}{maxSeq + 1:D3}";
                }

                model.AccNo = newAccNo;

                // Adjust balance for DR/CR
                decimal actualOpenBal = model.OpeningBalanceType == "DR" ? model.OpeningBalance : -model.OpeningBalance;

                // 1. Insert into Accounts
                string sqlAccount = @"
                    INSERT INTO Accounts (AccNo, AccTitle, SubAccOf, Type, openbal, opendate, Balance, Parent, Active, UserName, ComputerName)
                    VALUES (@AccNo, @VendorName, @ParentAccNo, @Type, @OpenBal, @OpenDate, @OpenBal, 0, @Active, 'Developer', 'Developer');
                ";
                
                var accParams = new {
                    AccNo = model.AccNo,
                    VendorName = model.VendorName,
                    ParentAccNo = model.ParentAccNo,
                    Type = accType,
                    OpenBal = actualOpenBal,
                    OpenDate = DateTime.Now.Date,
                    Active = model.Inactive ? 0 : 1
                };
                
                await connection.ExecuteAsync(sqlAccount, accParams, transaction);

                // 2. Insert into Venders
                string sqlVendor = @"
                    INSERT INTO Venders (
                        AccNo, Phone1, Phone2, Mobile, CPhone, CEmail, Fax1, Fax2, Address, 
                        ContactPerson, ImportVender, MakerNo, BankAccNo, ProcessID, 
                        VenderNameUrdu, VenderDescription
                    ) VALUES (
                        @AccNo, @Phone1, @Phone2, @Mobile, @CPhone, @CEmail, @Fax1, @Fax2, @Address, 
                        @ContactPerson, @ImportVendor, @MakerNo, @BankAccNo, @ProcessID, 
                        @VendorNameUrdu, @VendorDescription
                    );
                    SELECT CAST(SCOPE_IDENTITY() AS INT);
                ";
                
                int newVendId = await connection.ExecuteScalarAsync<int>(sqlVendor, model, transaction);
                model.VendID = newVendId;

                // TODO: Save pics if required (currently handled via separate Base64 logic, can be added later if needed)

                transaction.Commit();
                return newVendId;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                _logger.LogError(ex, "Error inserting new Vendor/Account");
                throw;
            }
        }

        public async Task UpdateVendorAsync(NewVendorViewModel model)
        {
            using IDbConnection connection = new SqlConnection(_connectionString);
            connection.Open();
            using IDbTransaction transaction = connection.BeginTransaction();

            try
            {
                decimal actualOpenBal = model.OpeningBalanceType == "DR" ? model.OpeningBalance : -model.OpeningBalance;

                // 1. Update Accounts
                // Get old OpenBal to adjust Balance
                decimal oldOpenBal = await connection.ExecuteScalarAsync<decimal>(
                    "SELECT openbal FROM Accounts WHERE AccNo = @AccNo", new { AccNo = model.AccNo }, transaction);
                
                decimal diff = actualOpenBal - oldOpenBal;

                string sqlAccount = @"
                    UPDATE Accounts 
                    SET AccTitle = @VendorName, 
                        openbal = @OpenBal, 
                        Balance = Balance + @Diff, 
                        Active = @Active
                    WHERE AccNo = @AccNo;
                ";
                
                var accParams = new {
                    VendorName = model.VendorName,
                    OpenBal = actualOpenBal,
                    Diff = diff,
                    Active = model.Inactive ? 0 : 1,
                    AccNo = model.AccNo
                };
                
                await connection.ExecuteAsync(sqlAccount, accParams, transaction);

                // 2. Update Venders
                string sqlVendor = @"
                    UPDATE Venders SET 
                        Phone1 = @Phone1, 
                        Phone2 = @Phone2, 
                        Mobile = @Mobile, 
                        CPhone = @CPhone, 
                        CEmail = @CEmail, 
                        Fax1 = @Fax1, 
                        Fax2 = @Fax2, 
                        Address = @Address, 
                        ContactPerson = @ContactPerson, 
                        ImportVender = @ImportVendor, 
                        MakerNo = @MakerNo, 
                        BankAccNo = @BankAccNo, 
                        ProcessID = @ProcessID, 
                        VenderNameUrdu = @VendorNameUrdu, 
                        VenderDescription = @VendorDescription
                    WHERE VendID = @VendID;
                ";
                
                await connection.ExecuteAsync(sqlVendor, model, transaction);

                transaction.Commit();
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                _logger.LogError(ex, "Error updating Vendor/Account");
                throw;
            }
        }

        public async Task DeleteVendorAsync(int id)
        {
            // Usually deleting a vendor also implies deleting the account if it has no transactions.
            // Due to foreign key cascading (ON DELETE CASCADE), deleting the Account will delete the Vendor.
            // Or we delete Vendor and then Account.
            using IDbConnection connection = new SqlConnection(_connectionString);
            
            string accNo = await connection.ExecuteScalarAsync<string>("SELECT AccNo FROM Venders WHERE VendID = @Id", new { Id = id });
            
            if (!string.IsNullOrEmpty(accNo))
            {
                // Deleting the Account should cascade and delete the Vendor, or at least clean up both
                await connection.ExecuteAsync("DELETE FROM Accounts WHERE AccNo = @AccNo", new { AccNo = accNo });
            }
        }

        public async Task ToggleVendorStatusAsync(string accNo, bool makeActive)
        {
            using IDbConnection connection = new SqlConnection(_connectionString);
            string sql = "UPDATE Accounts SET Active = @Active WHERE AccNo = @AccNo";
            await connection.ExecuteAsync(sql, new { Active = makeActive ? 1 : 0, AccNo = accNo });
        }
    }
}
