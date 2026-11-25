using Dapper;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Production
{
    public class ProductionItemListDataAccess : IProductionItemListDataAccess
    {
        private readonly IConfiguration _config;

        public ProductionItemListDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection string is not configured.");

        // ─────────────────────────────────────────────────────────────────────
        // MAIN LIST QUERY — mirrors VB6 getCondition() + RefreshLV() logic
        // ─────────────────────────────────────────────────────────────────────
        public async Task<List<ProductionItemListItem>> GetListAsync(ProductionItemListFilters f)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            DateTime dtFrom = f.DtFrom;
            DateTime dtTo = f.DtTo;
            if (f.DateRangeMode == 1) { dtFrom = DateTime.Today; dtTo = DateTime.Today.AddDays(15); }
            else if (f.DateRangeMode == 2) { dtFrom = DateTime.Today; dtTo = DateTime.Today.AddDays(30); }
            else if (f.DateRangeMode == 3) { dtFrom = DateTime.Today; dtTo = DateTime.Today.AddDays(45); }

            var p = new DynamicParameters();
            p.Add("LotNo", string.IsNullOrWhiteSpace(f.LotNo) ? null : f.LotNo.Trim());
            p.Add("DateRangeMode", f.DateRangeMode);
            p.Add("DtFrom", f.DateRangeMode != 0 ? dtFrom : (DateTime?)null);
            p.Add("DtTo", f.DateRangeMode != 0 ? dtTo : (DateTime?)null);
            p.Add("CustCode", string.IsNullOrEmpty(f.CustCode) || f.CustCode == "0" ? null : f.CustCode);
            p.Add("Country", string.IsNullOrEmpty(f.Country) || f.Country == "0" ? null : f.Country);
            p.Add("GroupID", int.TryParse(f.ItemGroupId, out int gid) ? gid : 0);
            p.Add("OrderNo", string.IsNullOrEmpty(f.OrderNo) || f.OrderNo == "0" ? null : f.OrderNo);
            p.Add("ProcessID", int.TryParse(f.ProcessId, out int pid) ? pid : 0);
            p.Add("VendID", int.TryParse(f.MakerId, out int vid) ? vid : 0);
            p.Add("ItemCode", string.IsNullOrEmpty(f.ItemCode) || f.ItemCode == "0" ? null : f.ItemCode);
            p.Add("ShowRepairLots", f.ShowRepairLots);
            p.Add("ShowRegularLots", f.ShowRegularLots);
            p.Add("ShowLotsOnly", f.ShowLotsOnly);
            p.Add("ShowNotYetStarted", f.ShowNotYetStarted);
            p.Add("Hub", string.IsNullOrEmpty(f.Hub) || f.Hub == "0" ? null : f.Hub);

            var results = await db.QueryAsync<ProductionItemListItem>(
                "SP_GetProductionItemList",
                p,
                commandType: CommandType.StoredProcedure,
                commandTimeout: 120);

            return results.ToList();
        }

        // ─────────────────────────────────────────────────────────────────────
        // LOOKUP QUERIES
        // ─────────────────────────────────────────────────────────────────────
        public async Task<List<LookupItemString>> GetCustomersAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT CustCode AS Id, CustCode AS Name FROM ForeignCustomers ORDER BY CustCode";
            return (await db.QueryAsync<LookupItemString>(sql)).ToList();
        }

        public async Task<List<LookupItemString>> GetCountriesByCustomerAsync(string custCode)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql;
            object param;
            if (string.IsNullOrEmpty(custCode) || custCode == "0")
            {
                sql = "SELECT DISTINCT Country AS Id, Country AS Name FROM ForeignCustomers WHERE Country IS NOT NULL AND Country <> '' ORDER BY Country";
                param = new { };
            }
            else
            {
                sql = "SELECT DISTINCT Country AS Id, Country AS Name FROM ForeignCustomers WHERE CustCode = @CustCode AND Country IS NOT NULL AND Country <> '' ORDER BY Country";
                param = new { CustCode = custCode };
            }
            return (await db.QueryAsync<LookupItemString>(sql, param)).ToList();
        }

        public async Task<List<LookupItemString>> GetOrdersByCustomerAsync(string custCode)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql;
            object param;
            if (string.IsNullOrEmpty(custCode) || custCode == "0")
            {
                sql = "SELECT OrderNo AS Id, OrderNo + ' {' + ISNULL(InternalRefNo,'') + '}' AS Name FROM FCustomerOrders ORDER BY OrderNo";
                param = new { };
            }
            else
            {
                sql = "SELECT OrderNo AS Id, OrderNo + ' {' + ISNULL(InternalRefNo,'') + '}' AS Name FROM FCustomerOrders WHERE CustCode = @CustCode ORDER BY OrderNo";
                param = new { CustCode = custCode };
            }
            return (await db.QueryAsync<LookupItemString>(sql, param)).ToList();
        }

        public async Task<List<LookupItemInt>> GetItemGroupsAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT ID AS Id, Description AS Name FROM ItemGroups ORDER BY Description";
            return (await db.QueryAsync<LookupItemInt>(sql)).ToList();
        }

        public async Task<List<LookupItemInt>> GetProcessesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT ProcessID AS Id, Description AS Name FROM Processes ORDER BY SNo";
            return (await db.QueryAsync<LookupItemInt>(sql)).ToList();
        }

        public async Task<List<LookupItemInt>> GetMakersAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT VendID AS Id, VendID1 + ' ' + VenderName AS Name FROM Makers ORDER BY VendID";
            return (await db.QueryAsync<LookupItemInt>(sql)).ToList();
        }

        public async Task<List<LookupItemString>> GetItemsAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT ItemID AS Id, ItemID + ' ' + ItemName AS Name FROM Items ORDER BY ItemID";
            return (await db.QueryAsync<LookupItemString>(sql)).ToList();
        }

        public async Task<List<string>> GetHubsAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT DISTINCT Hub_Name FROM ProcessGroupsProcesses WHERE Hub_Name IS NOT NULL AND Hub_Name <> '' ORDER BY Hub_Name";
            return (await db.QueryAsync<string>(sql)).ToList();
        }

        // ─────────────────────────────────────────────────────────────────────
        // REMARKS UPSERT — mirrors VB6 txtEdit KeyPress logic
        // POI_Remarks stores per LotNo (EntryType=0) or per OrderNo+ItemCode (EntryType=1)
        // ─────────────────────────────────────────────────────────────────────
        public async Task<bool> SaveRemarksAsync(string lotNo, int entryType, string orderNo, string itemCode, string remarks)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                IF @EntryType = 0
                BEGIN
                    IF EXISTS (SELECT 1 FROM POI_Remarks WHERE LotNo = @LotNo)
                        UPDATE POI_Remarks SET Remarks = @Remarks, EntryType = 0 WHERE LotNo = @LotNo
                    ELSE
                        INSERT INTO POI_Remarks(EntryType, LotNo, OrderNo, ItemCode, Remarks)
                        VALUES (0, @LotNo, @OrderNo, @ItemCode, @Remarks)
                END
                ELSE
                BEGIN
                    IF EXISTS (SELECT 1 FROM POI_Remarks WHERE ItemCode = @ItemCode AND OrderNo = @OrderNo)
                        UPDATE POI_Remarks SET Remarks = @Remarks, EntryType = 1 WHERE ItemCode = @ItemCode AND OrderNo = @OrderNo
                    ELSE
                        INSERT INTO POI_Remarks(EntryType, LotNo, OrderNo, ItemCode, Remarks)
                        VALUES (1, @LotNo, @OrderNo, @ItemCode, @Remarks)
                END";
            int rows = await db.ExecuteAsync(sql, new { LotNo = lotNo, EntryType = entryType, OrderNo = orderNo, ItemCode = itemCode, Remarks = remarks });
            return rows > 0;
        }
    }
}
