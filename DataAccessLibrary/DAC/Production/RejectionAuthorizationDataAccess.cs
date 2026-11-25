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
    public class RejectionAuthorizationDataAccess : IRejectionAuthorizationDataAccess
    {
        private readonly IConfiguration _config;

        public RejectionAuthorizationDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection string is not configured.");

        public async Task<int> GetFactoryMakerIdAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT TOP 1 DataValue FROM GeneralData WHERE DataName = 'FactoryMaker'";
            string? val = await db.QueryFirstOrDefaultAsync<string?>(sql);
            return int.TryParse(val, out int id) ? id : 0;
        }

        public async Task<List<RejectionAuthMakerLookup>> GetMakersAsync(int factoryMakerId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT VendID, ISNULL(VendID1, '') AS VendID1, ISNULL(VenderName, '') AS VenderName
                FROM Makers
                WHERE VendID <> @FactoryMakerId
                ORDER BY VenderName";
            return (await db.QueryAsync<RejectionAuthMakerLookup>(sql, new { FactoryMakerId = factoryMakerId })).ToList();
        }

        public async Task<List<RejectionAuthEmployeeLookup>> GetEmployeesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT EmpID, ISNULL(Name, '') AS Name, ISNULL(FName, '') AS FName, ISNULL(Designation, '') AS Designation
                FROM Employees
                ORDER BY Name";
            return (await db.QueryAsync<RejectionAuthEmployeeLookup>(sql)).ToList();
        }

        public async Task<List<RejectionAuthCategoryLookup>> GetCategoriesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT CatID, ISNULL(Description, '') AS Description
                FROM ItemCatagories
                ORDER BY Description";
            return (await db.QueryAsync<RejectionAuthCategoryLookup>(sql)).ToList();
        }

        public async Task<List<RejectionAuthItemModel>> GetPendingRejectionsAsync(RejectionAuthFilterModel filter)
        {
            return await QueryVMakerRepairAsync(filter, isHistory: false);
        }

        public async Task<List<RejectionAuthItemModel>> GetHistoryAsync(RejectionAuthFilterModel filter)
        {
            return await QueryVMakerRepairAsync(filter, isHistory: true);
        }

        private async Task<List<RejectionAuthItemModel>> QueryVMakerRepairAsync(RejectionAuthFilterModel filter, bool isHistory)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            var parameters = new DynamicParameters();
            parameters.Add("@DtFrom", filter.DtFrom.Date);
            parameters.Add("@DtTo", filter.DtTo.Date);

            var whereParts = new List<string>
            {
                "ShowInBilling = 1",
                isHistory ? "DeductFromBilling IS NOT NULL" : "VchrNo IS NULL AND DeductFromBilling IS NULL",
                "CAST(DT AS DATE) BETWEEN @DtFrom AND @DtTo"
            };

            if (filter.ShowEmployees)
            {
                if (!string.IsNullOrWhiteSpace(filter.EmpID) && filter.EmpID != "0")
                {
                    whereParts.Add("EmpID = @EmpID");
                    parameters.Add("@EmpID", filter.EmpID);
                }
            }
            else if (filter.VendID > 0)
            {
                whereParts.Add("VendID = @VendID");
                parameters.Add("@VendID", filter.VendID);
            }

            if (filter.CategoryIds != null && filter.CategoryIds.Any())
            {
                whereParts.Add("CatID IN @CategoryIds");
                parameters.Add("@CategoryIds", filter.CategoryIds);
            }

            if (!string.IsNullOrWhiteSpace(filter.LotNo))
            {
                whereParts.Add("LotNo = @LotNo");
                parameters.Add("@LotNo", filter.LotNo.Trim());
            }

            string sql = $@"
                SELECT EntryID,
                       ISNULL(RecieptID, '') AS RecieptID,
                       DT,
                       ISNULL(ItemCode, '') AS ItemCode,
                       ISNULL(ItemName, '') AS ItemName,
                       ISNULL(Description, '') AS Description,
                       ISNULL(Qty, 0) AS Qty,
                       ISNULL(Rate, 0) AS Rate,
                       ISNULL(RepairAmt, 0) AS RepairAmt,
                       ISNULL(LotNo, '') AS LotNo,
                       ISNULL(RcvdQty, 0) AS RcvdQty,
                       VendID,
                       FromVendID,
                       ISNULL(VendID1, '') AS VendID1,
                       ISNULL(VenderName, '') AS VenderName,
                       EmpID,
                       CatID,
                       DeductFromBilling
                FROM VMakerRepair
                WHERE {string.Join(" AND ", whereParts)}
                ORDER BY DT DESC, EntryID DESC";

            return (await db.QueryAsync<RejectionAuthItemModel>(sql, parameters)).ToList();
        }

        public async Task<int> UpdateDeductFromBillingAsync(List<int> entryIds, bool deduct)
        {
            if (entryIds == null || !entryIds.Any())
                return 0;

            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "UPDATE MakerRepair SET DeductFromBilling = @Deduct WHERE EntryID IN @EntryIds";
            return await db.ExecuteAsync(sql, new { Deduct = deduct ? 1 : 0, EntryIds = entryIds });
        }

        public async Task<bool> UpdateRateAsync(int entryId, int rate)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "UPDATE MakerRepair SET Rate = @Rate WHERE EntryID = @EntryID";
            int affected = await db.ExecuteAsync(sql, new { EntryID = entryId, Rate = rate });
            return affected > 0;
        }
    }
}
