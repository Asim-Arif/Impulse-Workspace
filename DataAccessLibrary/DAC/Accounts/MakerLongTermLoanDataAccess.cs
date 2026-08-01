using Dapper;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Accounts
{
    public class MakerLongTermLoanDataAccess : IMakerLongTermLoanDataAccess
    {
        private readonly IConfiguration _config;

        public MakerLongTermLoanDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection");

        public async Task<List<MakerAccountModel>> GetMakerAccountsAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT VendID, AccNo, ISNULL(MakerNo, '') AS MakerNo, AccTitle AS MakerName, AccTitle
                           FROM VVendersAndMakers 
                           WHERE AccNo IS NOT NULL AND AccNo <> '' 
                           ORDER BY AccTitle";

            var list = (await db.QueryAsync<MakerAccountModel>(sql)).ToList();
            return list;
        }

        public async Task<decimal> GetMakerLongTermBalanceAsync(string accNo)
        {
            if (string.IsNullOrWhiteSpace(accNo)) return 0m;

            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT ISNULL(SUM(Amount - AmountCleared), 0) FROM MakerAdvances WHERE AccNo = @AccNo";

            return await db.ExecuteScalarAsync<decimal>(sql, new { AccNo = accNo });
        }

        public async Task<List<MakerMasterPoModel>> GetMakerMasterPOsAsync(long vendId, string accNo)
        {
            if (string.IsNullOrWhiteSpace(accNo)) return new List<MakerMasterPoModel>();

            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT DISTINCT MasterPONo FROM VendIssued WHERE (VendID = @VendId OR VendID = (SELECT VendID FROM Makers WHERE AccNo = @AccNo)) AND MasterPONo IS NOT NULL AND MasterPONo <> ''
                           UNION
                           SELECT DISTINCT OrderNo AS MasterPONo FROM VendOrders WHERE (VendID = @AccNo OR VendID = CAST(@VendId AS nvarchar)) AND OrderNo IS NOT NULL AND OrderNo <> ''";

            var pos = (await db.QueryAsync<string>(sql, new { VendId = vendId, AccNo = accNo })).ToList();

            var list = new List<MakerMasterPoModel>();
            foreach (var po in pos)
            {
                list.Add(new MakerMasterPoModel
                {
                    MasterPONo = po,
                    POValue = 0,
                    LTTaken = 0,
                    STTaken = 0
                });
            }
            return list;
        }

        public async Task<MakerMasterPoModel> GetMasterPoDetailsAsync(string masterPoNo, bool isVendor)
        {
            if (string.IsNullOrWhiteSpace(masterPoNo)) return new MakerMasterPoModel();

            using IDbConnection db = new SqlConnection(ConnectionString);

            string poValueSql = isVendor
                ? @"SELECT ISNULL(SUM(Qty * Rate), 0) FROM VendOrderDetail WHERE OrderNo = @MasterPONo"
                : @"SELECT ISNULL(SUM(vd.IssQty * vd.Rate), 0) FROM VendIssdDetail vd INNER JOIN VendIssued vi ON vi.EntryID = vd.RefID WHERE vi.MasterPONo = @MasterPONo";

            decimal poValue = await db.ExecuteScalarAsync<decimal>(poValueSql, new { MasterPONo = masterPoNo });

            string ltSql = @"SELECT ISNULL(SUM(Amount), 0) FROM MakerAdvances WHERE MasterPONo = @MasterPONo";
            decimal ltTaken = await db.ExecuteScalarAsync<decimal>(ltSql, new { MasterPONo = masterPoNo });

            string stSql = @"SELECT ISNULL(SUM(Amount), 0) FROM MakerAdvancesShort WHERE MasterPONo = @MasterPONo";
            decimal stTaken = await db.ExecuteScalarAsync<decimal>(stSql, new { MasterPONo = masterPoNo });

            return new MakerMasterPoModel
            {
                MasterPONo = masterPoNo,
                POValue = poValue,
                LTTaken = ltTaken,
                STTaken = stTaken
            };
        }

        public async Task SaveMakerLongTermLoanAsync(MakerLongTermLoanDto dto, string vchrNo)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"INSERT INTO MakerAdvances(AccNo, DT, Type, Description, Amount, DAmount, AmountCleared, AccVoucherNo, MasterPONo)
                           VALUES(@AccNo, @DT, 1, @Description, @Amount, @DAmount, 0, @AccVoucherNo, @MasterPONo)";

            await db.ExecuteAsync(sql, new
            {
                AccNo = dto.AccNo,
                DT = dto.IssueDate,
                Description = dto.Description,
                Amount = dto.Amount,
                DAmount = dto.DAmount,
                AccVoucherNo = vchrNo,
                MasterPONo = dto.MasterPONo
            });
        }
    }
}
