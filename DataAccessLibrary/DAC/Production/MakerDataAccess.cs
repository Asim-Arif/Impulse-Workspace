using Dapper;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Production
{
    public class MakerDataAccess : IMakerDataAccess
    {
        private readonly IConfiguration _config;

        public MakerDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection") 
            ?? throw new InvalidOperationException("DefaultConnection configuration string is missing.");

        public async Task<MakerDto?> GetMakerByIdAsync(long vendId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT 
                            VendID, VendID1, VenderName, CompanyName, Maker_Second_Name, AccNo, SubAccOf,
                            Phone1, Phone2, Fax1 AS Phone3, Mobile, NICNo, NTNNo, Address, BankAccNo,
                            MaxLimit, TimeLimit,
                            AuthRequired, ShowRateOnPO, MakerNameUrdu, CNIC_PDF_FileName,
                            PaymentTerms, RefBy, RefByFName, RefByPhone1, RefByPhone2, RefByAddress,
                            ExcessQtyPercentage, MaximumRcvingsAgainstPO, MakerCapacity,
                            Active
                           FROM VMakers
                           WHERE VendID = @VendId";

            return await db.QueryFirstOrDefaultAsync<MakerDto>(sql, new { VendId = vendId });
        }

        public async Task<List<ProcessModel>> GetAllProcessesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT ProcessID, Description, Code, SNo 
                           FROM Processes 
                           ORDER BY SNo";

            var list = (await db.QueryAsync<ProcessModel>(sql)).ToList();
            return list;
        }

        public async Task<List<int>> GetMakerProcessIdsAsync(long vendId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT ProcessID FROM MakerProcesses WHERE VendID = @VendId";
            return (await db.QueryAsync<int>(sql, new { VendId = vendId })).ToList();
        }

        public async Task<List<MakerCategoryModel>> GetMakerCategoriesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT AccNo, AccTitle 
                           FROM Accounts 
                           WHERE Parent = 1 
                           ORDER BY AccTitle";

            return (await db.QueryAsync<MakerCategoryModel>(sql)).ToList();
        }

        public async Task<List<string>> GetPaymentTermsAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT DISTINCT PaymentTerms 
                           FROM Makers 
                           WHERE PaymentTerms IS NOT NULL AND PaymentTerms <> '' 
                           ORDER BY PaymentTerms";

            return (await db.QueryAsync<string>(sql)).ToList();
        }

        public async Task<string> GenerateNextMakerCodeAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT ISNULL(MAX(CAST(RIGHT(VendID1, 4) AS INT)), 0) + 1 
                           FROM Makers 
                           WHERE ISNUMERIC(RIGHT(VendID1, 4)) = 1";

            int nextNum = await db.ExecuteScalarAsync<int>(sql);
            if (nextNum == 0) nextNum = 1;
            return $"CONT-{nextNum:D4}";
        }

        public async Task<string> GenerateNextAccNoAsync(string parentAccNo)
        {
            if (string.IsNullOrWhiteSpace(parentAccNo)) return string.Empty;

            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT ISNULL(MAX(CAST(RIGHT(AccNo, 3) AS INT)), 0) + 1 
                           FROM Accounts 
                           WHERE SubAccOf = @ParentAccNo";

            int nextVal = await db.ExecuteScalarAsync<int>(sql, new { ParentAccNo = parentAccNo });
            return $"{parentAccNo}-{nextVal:D3}";
        }

        public async Task<bool> IsNicDuplicateAsync(string nicNo, long? excludeVendId)
        {
            if (string.IsNullOrWhiteSpace(nicNo)) return false;

            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT COUNT(1) FROM Makers WHERE NICNo = @NicNo";
            if (excludeVendId.HasValue && excludeVendId.Value > 0)
            {
                sql += " AND VendID <> @ExcludeVendId";
            }

            int count = await db.ExecuteScalarAsync<int>(sql, new { NicNo = nicNo, ExcludeVendId = excludeVendId });
            return count > 0;
        }

        public async Task<long> SaveMakerAsync(MakerDto dto)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                // 1. Insert into Accounts table
                string typePrefix = !string.IsNullOrEmpty(dto.SubAccOf) && dto.SubAccOf.Length >= 2 
                    ? dto.SubAccOf.Substring(0, 2) 
                    : "24";

                string accSql = @"INSERT INTO Accounts (AccNo, AccTitle, SubAccOf, Type, OpenBal, opendate, Balance, Active)
                                  VALUES (@AccNo, @AccTitle, @SubAccOf, @Type, 0, GETDATE(), 0, 1)";

                await db.ExecuteAsync(accSql, new
                {
                    AccNo = dto.AccNo,
                    AccTitle = dto.VenderName,
                    SubAccOf = dto.SubAccOf,
                    Type = typePrefix
                }, trans);

                // 2. Insert into Makers table
                string makerSql = @"INSERT INTO Makers 
                                    (VendID1, VenderName, CompanyName, Maker_Second_Name, Phone1, Phone2, Fax1, Mobile, NICNo, NTNNo, Address, BankAccNo,
                                     MaxLimit, TimeLimit, AuthRequired, ShowRateOnPO, MakerNameUrdu, CNIC_PDF_FileName,
                                     PaymentTerms, RefBy, RefByFName, RefByPhone1, RefByPhone2, RefByAddress,
                                     ExcessQtyPercentage, MaximumRcvingsAgainstPO, MakerCapacity, AccNo, Active)
                                    VALUES 
                                    (@VendID1, @VenderName, @CompanyName, @Maker_Second_Name, @Phone1, @Phone2, @Phone3, @Mobile, @NICNo, @NTNNo, @Address, @BankAccNo,
                                     @MaxLimit, @TimeLimit, @AuthRequired, @ShowRateOnPO, @MakerNameUrdu, @CNIC_PDF_FileName,
                                     @PaymentTerms, @RefBy, @RefByFName, @RefByPhone1, @RefByPhone2, @RefByAddress,
                                     @ExcessQtyPercentage, @MaximumRcvingsAgainstPO, @MakerCapacity, @AccNo, @Active)";

                await db.ExecuteAsync(makerSql, dto, trans);

                // 3. Get inserted VendID
                long newVendId = await db.ExecuteScalarAsync<long>("SELECT MAX(VendID) FROM Makers", null, trans);

                trans.Commit();
                return newVendId;
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }

        public async Task UpdateMakerAsync(MakerDto dto)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                // 1. Update Accounts table Title
                string accSql = @"UPDATE Accounts SET AccTitle = @VenderName WHERE AccNo = @AccNo";
                await db.ExecuteAsync(accSql, new { VenderName = dto.VenderName, AccNo = dto.AccNo }, trans);

                // 2. Update Makers table
                string makerSql = @"UPDATE Makers SET 
                                    VenderName = @VenderName,
                                    CompanyName = @CompanyName,
                                    Maker_Second_Name = @Maker_Second_Name,
                                    Phone1 = @Phone1,
                                    Phone2 = @Phone2,
                                    Fax1 = @Phone3,
                                    Mobile = @Mobile,
                                    NICNo = @NICNo,
                                    NTNNo = @NTNNo,
                                    Address = @Address,
                                    BankAccNo = @BankAccNo,
                                    MaxLimit = @MaxLimit,
                                    TimeLimit = @TimeLimit,
                                    AuthRequired = @AuthRequired,
                                    ShowRateOnPO = @ShowRateOnPO,
                                    MakerNameUrdu = @MakerNameUrdu,
                                    CNIC_PDF_FileName = @CNIC_PDF_FileName,
                                    PaymentTerms = @PaymentTerms,
                                    RefBy = @RefBy,
                                    RefByFName = @RefByFName,
                                    RefByPhone1 = @RefByPhone1,
                                    RefByPhone2 = @RefByPhone2,
                                    RefByAddress = @RefByAddress,
                                    ExcessQtyPercentage = @ExcessQtyPercentage,
                                    MaximumRcvingsAgainstPO = @MaximumRcvingsAgainstPO,
                                    MakerCapacity = @MakerCapacity,
                                    Active = @Active
                                    WHERE VendID = @VendID";

                await db.ExecuteAsync(makerSql, dto, trans);

                trans.Commit();
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }

        public async Task SaveMakerProcessesAsync(long vendId, List<int> processIds)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                await db.ExecuteAsync("DELETE FROM MakerProcesses WHERE VendID = @VendId", new { VendId = vendId }, trans);

                if (processIds != null && processIds.Any())
                {
                    string insertSql = "INSERT INTO MakerProcesses (VendID, ProcessID) VALUES (@VendId, @ProcessID)";
                    var batch = processIds.Select(pId => new { VendId = vendId, ProcessID = pId });
                    await db.ExecuteAsync(insertSql, batch, trans);
                }

                trans.Commit();
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }
    }
}
