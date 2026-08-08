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

            // If a specific lot number is entered, it overrides all other filters (mirrors legacy txtLotNo override)
            string whereClause;
            var p = new DynamicParameters();

            if (!string.IsNullOrWhiteSpace(f.LotNo))
            {
                whereClause = "WHERE LotNo = @LotNo";
                p.Add("LotNo", f.LotNo.Trim());
            }
            else
            {
                var sb = new StringBuilder("WHERE 1=1");

                // Date range (only applied when mode != 0 = All)
                if (f.DateRangeMode != 0)
                {
                    DateTime dtFrom = f.DtFrom;
                    DateTime dtTo = f.DtTo;
                    if (f.DateRangeMode == 1) { dtFrom = DateTime.Today; dtTo = DateTime.Today.AddDays(15); }
                    else if (f.DateRangeMode == 2) { dtFrom = DateTime.Today; dtTo = DateTime.Today.AddDays(30); }
                    else if (f.DateRangeMode == 3) { dtFrom = DateTime.Today; dtTo = DateTime.Today.AddDays(45); }
                    // mode 4 = Custom — use f.DtFrom / f.DtTo as-is

                    sb.Append(" AND DT BETWEEN @DtFrom AND @DtTo");
                    p.Add("DtFrom", dtFrom);
                    p.Add("DtTo", dtTo);
                }

                // Customer
                if (!string.IsNullOrEmpty(f.CustCode) && f.CustCode != "0")
                {
                    sb.Append(" AND CustCode = @CustCode");
                    p.Add("CustCode", f.CustCode);
                }

                // Country
                if (!string.IsNullOrEmpty(f.Country) && f.Country != "0")
                {
                    sb.Append(" AND Country = @Country");
                    p.Add("Country", f.Country);
                }

                // Item Group
                if (!string.IsNullOrEmpty(f.ItemGroupId) && f.ItemGroupId != "0")
                {
                    sb.Append(" AND GroupID = @GroupID");
                    p.Add("GroupID", int.Parse(f.ItemGroupId));
                }

                // Order No
                if (!string.IsNullOrEmpty(f.OrderNo) && f.OrderNo != "0")
                {
                    sb.Append(" AND OrderNo = @OrderNo");
                    p.Add("OrderNo", f.OrderNo);
                }

                // Process
                if (!string.IsNullOrEmpty(f.ProcessId) && f.ProcessId != "0")
                {
                    sb.Append(" AND ProcessID = @ProcessID");
                    p.Add("ProcessID", int.Parse(f.ProcessId));
                }

                // Maker (VendID)
                if (!string.IsNullOrEmpty(f.MakerId) && f.MakerId != "0")
                {
                    sb.Append(" AND VendID = @VendID");
                    p.Add("VendID", int.Parse(f.MakerId));
                }

                // Item Code
                if (!string.IsNullOrEmpty(f.ItemCode) && f.ItemCode != "0")
                {
                    sb.Append(" AND ItemCode = @ItemCode");
                    p.Add("ItemCode", f.ItemCode);
                }

                // Repair Lots toggle (ReWorkLot=1 means repair)
                if (!f.ShowRepairLots)
                    sb.Append(" AND ReWorkLot = 0");
                if (!f.ShowRegularLots)
                    sb.Append(" AND ReWorkLot = 1");

                // Lots Only toggle (EntryType=0 means lot, EntryType=1 means not-yet-started order item)
                if (!f.ShowLotsOnly)
                    sb.Append(" AND EntryType <> 0");
                if (!f.ShowNotYetStarted)
                    sb.Append(" AND EntryType <> 1");

                // Hub
                if (!string.IsNullOrEmpty(f.Hub) && f.Hub != "0")
                {
                    sb.Append(" AND Hub_Name = @Hub");
                    p.Add("Hub", f.Hub);
                }

                // Always filter out zero-qty rows
                sb.Append(" AND Qty > 0");

                whereClause = sb.ToString();
            }

            string sql = $"SELECT * FROM VRunningLots_POI WITH (NOLOCK) {whereClause} ORDER BY StartDT DESC";
            var results = await db.QueryAsync<ProductionItemListItem>(sql, p, commandTimeout: 120);
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
            const string sql = "SELECT ItemID AS Id, ItemID + ' ' + ItemName AS Name FROM VItems ORDER BY ItemID";
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
                IF EXISTS (SELECT 1 FROM POI_Remarks WHERE LotNo = @LotNo AND EntryType = @EntryType)
                    UPDATE POI_Remarks SET Remarks = @Remarks WHERE LotNo = @LotNo AND EntryType = @EntryType
                ELSE
                    INSERT INTO POI_Remarks(EntryType, LotNo, OrderNo, ItemCode, Remarks)
                    VALUES (@EntryType, @LotNo, @OrderNo, @ItemCode, @Remarks)";
            int rows = await db.ExecuteAsync(sql, new { LotNo = lotNo, EntryType = entryType, OrderNo = orderNo, ItemCode = itemCode, Remarks = remarks });
            return rows > 0;
        }
    }
}
