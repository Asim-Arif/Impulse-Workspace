using Dapper;
using DataAccessLibrary.Interface.Production;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System;
using System.Data;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Production
{
    public class ProductionSequenceDataAccess : IProductionSequenceDataAccess
    {
        private readonly IConfiguration _config;

        public ProductionSequenceDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection configuration string is missing.");

        public async Task<string> GetNextBatchNoAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT MAX(CAST(RIGHT(Batch_No, LEN(Batch_No) - 3) AS INT)) 
                           FROM VendIssdDetail 
                           WHERE Batch_No LIKE 'B17%' AND ISNUMERIC(RIGHT(Batch_No, LEN(Batch_No) - 3)) = 1";

            int maxNo = await db.ExecuteScalarAsync<int?>(sql) ?? 0;
            return $"B17{maxNo + 1}";
        }

        public async Task<string> GetNextMasterPONoAsync(DateTime dt)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string datePattern = dt.ToString("MMyy");

            string sql = @"SELECT MAX(CAST(LEFT(MasterPONo, CHARINDEX('/', MasterPONo) - 1) AS INT)) 
                           FROM VendIssued 
                           WHERE RIGHT(MasterPONo, 4) = @DatePattern AND MasterPONo IS NOT NULL AND CHARINDEX('/', MasterPONo) > 0";

            int maxNo = await db.ExecuteScalarAsync<int?>(sql, new { DatePattern = datePattern }) ?? 0;
            return $"{maxNo + 1}/{datePattern}";
        }

        public async Task<string> GetNextHeaderReceiptIDAsync(DateTime dt)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string dtFormatted = dt.ToString("ddMMyy");
            string dtSearchFormatted = dt.ToString("dd MMM yy");

            string sql = @"SELECT MAX(CAST(RIGHT(RecieptID, LEN(RecieptID) - 12) AS INT)) 
                           FROM VendIssued 
                           WHERE CONVERT(varchar, DT, 6) = @DtSearch AND RecieptID LIKE 'M-ISU-%'";

            int lastNo = await db.ExecuteScalarAsync<int?>(sql, new { DtSearch = dtSearchFormatted }) ?? 0;
            return $"M-ISU-{dtFormatted}{lastNo + 1}";
        }

        public async Task<string> GetNextDetailReceiptIDAsync(DateTime dt)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string dtFormatted = dt.ToString("ddMMyy");

            string sql = @"SELECT MAX(CAST(RIGHT(RecieptID, LEN(RecieptID) - 10) AS INT)) 
                           FROM VendIssdDetail 
                           WHERE RecieptID LIKE 'ISU-%'";

            int lastNo = await db.ExecuteScalarAsync<int?>(sql) ?? 0;
            return $"ISU-{dtFormatted}{lastNo + 1}";
        }

        public async Task<string> GetNextSubLotNoAsync(string refLotNo)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            if (string.IsNullOrWhiteSpace(refLotNo))
            {
                string maxSql = @"SELECT MAX(CAST(LotNo AS BIGINT)) FROM VendIssdDetail WHERE ISNUMERIC(LotNo) = 1";
                long maxLot = await db.ExecuteScalarAsync<long?>(maxSql) ?? 0;
                return (maxLot + 1).ToString();
            }

            string subSql = @"SELECT MAX(CAST(RIGHT(LotNo, LEN(LotNo) - LEN(@RefLotNo)) AS INT)) 
                             FROM VendIssdDetail 
                             WHERE LotNo LIKE @RefLotNo + '%' AND LEN(LotNo) > LEN(@RefLotNo) AND ISNUMERIC(RIGHT(LotNo, LEN(LotNo) - LEN(@RefLotNo))) = 1";

            int nextSubNo = (await db.ExecuteScalarAsync<int?>(subSql, new { RefLotNo = refLotNo }) ?? 0) + 1;
            return $"{refLotNo}{nextSubNo}";
        }

        public async Task<string> GetNextReceivingReceiptIDAsync(DateTime dt)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string dtFormatted = dt.ToString("ddMMyy");

            string sql = @"SELECT MAX(CAST(RIGHT(RecieptID, LEN(RecieptID) - 10) AS INT)) 
                           FROM VendRcvdDetail 
                           WHERE UPPER(LEFT(RecieptID, 3)) = 'RCV' AND LEN(RecieptID) >= 11";

            int lastNo = await db.ExecuteScalarAsync<int?>(sql) ?? 0;
            return $"RCV-{dtFormatted}{lastNo + 1}";
        }

        public async Task<string> GetNextMainLotNoAsync(DateTime dt)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string yearPrefix = dt.ToString("yy");

            string sql = @"
                SELECT MAX(SeqNo) FROM (
                    SELECT CAST(SUBSTRING(LotNo, 3, CASE WHEN CHARINDEX('-', LotNo) > 0 THEN CHARINDEX('-', LotNo) - 3 ELSE LEN(LotNo) END) AS INT) AS SeqNo
                    FROM Lots_List WHERE LEFT(LotNo, 2) = @YearPrefix AND LEN(LotNo) = 7 AND ISNUMERIC(SUBSTRING(LotNo, 3, 5)) = 1
                    UNION ALL
                    SELECT CAST(SUBSTRING(LotNo, 3, CASE WHEN CHARINDEX('-', LotNo) > 0 THEN CHARINDEX('-', LotNo) - 3 ELSE LEN(LotNo) END) AS INT) AS SeqNo
                    FROM VendRcvdDetail WHERE LEFT(LotNo, 2) = @YearPrefix AND LEN(LotNo) = 7 AND ISNUMERIC(SUBSTRING(LotNo, 3, 5)) = 1
                    UNION ALL
                    SELECT CAST(SUBSTRING(LotNo, 3, CASE WHEN CHARINDEX('-', LotNo) > 0 THEN CHARINDEX('-', LotNo) - 3 ELSE LEN(LotNo) END) AS INT) AS SeqNo
                    FROM VendIssdDetail WHERE LEFT(LotNo, 2) = @YearPrefix AND LEN(LotNo) = 7 AND ISNUMERIC(SUBSTRING(LotNo, 3, 5)) = 1
                    UNION ALL
                    SELECT CAST(SUBSTRING(LotNo_Manual, 3, CASE WHEN CHARINDEX('-', LotNo_Manual) > 0 THEN CHARINDEX('-', LotNo_Manual) - 3 ELSE LEN(LotNo_Manual) END) AS INT) AS SeqNo
                    FROM StockOrderOpening WHERE LEFT(LotNo_Manual, 2) = @YearPrefix AND LEN(LotNo_Manual) = 7 AND ISNUMERIC(LotNo_Manual) = 1
                ) T";

            int maxSeq = await db.ExecuteScalarAsync<int?>(sql, new { YearPrefix = yearPrefix }) ?? 0;
            if (maxSeq == 0)
            {
                maxSeq = 53000;
            }
            else
            {
                maxSeq++;
            }

            return $"{yearPrefix}{maxSeq:D5}";
        }
    }
}
