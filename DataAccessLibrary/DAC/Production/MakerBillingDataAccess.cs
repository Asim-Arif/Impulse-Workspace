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
    public class MakerBillingDataAccess : IMakerBillingDataAccess
    {
        private readonly IConfiguration _config;

        public MakerBillingDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? _config.GetConnectionString("ImpulseConnection")
            ?? string.Empty;

        public async Task<List<LookupItemInt>> GetMakersAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT 0 AS Id, '<Select Maker>' AS Name
                UNION ALL
                SELECT VendID AS Id, ISNULL('{ ' + VendID1 + ' } ' + VenderName, VenderName) AS Name
                FROM Makers
                WHERE ISNULL(Active, 1) = 1
                ORDER BY Name";
            return (await db.QueryAsync<LookupItemInt>(sql)).ToList();
        }

        public async Task<List<LookupItemInt>> GetCategoriesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT 0 AS Id, '<All Categories>' AS Name
                UNION ALL
                SELECT CatID AS Id, Description AS Name
                FROM ItemCatagories
                ORDER BY Name";
            return (await db.QueryAsync<LookupItemInt>(sql)).ToList();
        }

        public async Task<List<LookupItemString>> GetDebitHeadsAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT AccNo AS Id, ISNULL(AccTitle + ' { ' + AccNo + ' }', AccNo) AS Name
                FROM Accounts
                WHERE Active = 1 AND Parent = 0
                ORDER BY AccNo";
            return (await db.QueryAsync<LookupItemString>(sql)).ToList();
        }

        public async Task<string> GetMakerAccountNoAsync(int vendId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            return await db.ExecuteScalarAsync<string>("SELECT TOP 1 AccNo FROM Makers WHERE VendID = @VendID", new { VendID = vendId }) ?? string.Empty;
        }

        public async Task<string> GetMakerCodeAsync(int vendId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            return await db.ExecuteScalarAsync<string>("SELECT TOP 1 VendID1 FROM Makers WHERE VendID = @VendID", new { VendID = vendId }) ?? string.Empty;
        }

        public async Task<long> GetNextBillNoAsync(int vendId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            return await db.ExecuteScalarAsync<long>(@"
                SELECT ISNULL(MAX(BillNo), 0) + 1
                FROM MakerPostedBills
                WHERE BillNo IS NOT NULL AND VendID = @VendID", new { VendID = vendId });
        }

        public async Task<List<MakerBillingLotItem>> GetMakerBillingLotsAsync(MakerBillingFilter filter)
        {
            if (filter.VendID <= 0) return new List<MakerBillingLotItem>();

            using IDbConnection db = new SqlConnection(ConnectionString);

            var p = new DynamicParameters();
            p.Add("@MakerID", filter.VendID);
            p.Add("@DTFrom", filter.DtFrom.Date);
            p.Add("@DTTo", filter.DtTo.Date.AddDays(1).AddTicks(-1));
            p.Add("@CatID", string.IsNullOrWhiteSpace(filter.CatIDs) ? "0" : filter.CatIDs);
            p.Add("@ProcessID", 0);
            p.Add("@MasterPONo", string.IsNullOrWhiteSpace(filter.MasterPONo) ? string.Empty : filter.MasterPONo.Trim());

            var rawItems = (await db.QueryAsync<dynamic>("MakerBilling_RcvingWise_SP", p, commandType: CommandType.StoredProcedure)).ToList();

            // Fetch hidden lot IDs for this maker
            var hiddenIds = (await db.QueryAsync<long>(@"
                SELECT VRD_RefID FROM MakerBilling_Receivings_Hidden
                WHERE VRD_RefID IS NOT NULL")).ToHashSet();

            var items = new List<MakerBillingLotItem>();
            foreach (var r in rawItems)
            {
                var dict = (IDictionary<string, object>)r;

                long rcvDetailEntryId = dict.ContainsKey("RcvDetailEntryID") && dict["RcvDetailEntryID"] != null
                    ? Convert.ToInt64(dict["RcvDetailEntryID"]) : 0;

                if (hiddenIds.Contains(rcvDetailEntryId)) continue;

                long viEntryId = dict.ContainsKey("EntryID") && dict["EntryID"] != null
                    ? Convert.ToInt64(dict["EntryID"]) : 0;

                long issDetailRefId = dict.ContainsKey("Issue_RefID") && dict["Issue_RefID"] != null
                    ? Convert.ToInt64(dict["Issue_RefID"]) : 0;

                string rcvReceiptId = dict.ContainsKey("RcvRecieptID") && dict["RcvRecieptID"] != null
                    ? dict["RcvRecieptID"].ToString()! : string.Empty;

                DateTime dt = dict.ContainsKey("RcvDT") && dict["RcvDT"] != null
                    ? Convert.ToDateTime(dict["RcvDT"])
                    : (dict.ContainsKey("DT") && dict["DT"] != null ? Convert.ToDateTime(dict["DT"]) : DateTime.Today);

                string orderNo = dict.ContainsKey("OrderNo") && dict["OrderNo"] != null ? dict["OrderNo"].ToString()! : string.Empty;
                string itemCode = dict.ContainsKey("ItemID") && dict["ItemID"] != null ? dict["ItemID"].ToString()! : string.Empty;
                string itemName = dict.ContainsKey("ItemName") && dict["ItemName"] != null ? dict["ItemName"].ToString()! : string.Empty;
                decimal issQty = dict.ContainsKey("IssQty") && dict["IssQty"] != null ? Convert.ToDecimal(dict["IssQty"]) : 0;
                decimal rate = dict.ContainsKey("Rate") && dict["Rate"] != null ? Convert.ToDecimal(dict["Rate"]) : 0;
                decimal tempAdvance = dict.ContainsKey("TempAdvance") && dict["TempAdvance"] != null ? Convert.ToDecimal(dict["TempAdvance"]) : 0;
                decimal repairAmt = dict.ContainsKey("RepairAmt") && dict["RepairAmt"] != null ? Convert.ToDecimal(dict["RepairAmt"]) : 0;

                decimal rcvdValue = dict.ContainsKey("BillingProcAmount") && dict["BillingProcAmount"] != null
                    ? Convert.ToDecimal(dict["BillingProcAmount"])
                    : (dict.ContainsKey("RcvdValue") && dict["RcvdValue"] != null ? Convert.ToDecimal(dict["RcvdValue"]) : 0);

                decimal rcvdQty = dict.ContainsKey("BillingProcRcvdQty") && dict["BillingProcRcvdQty"] != null
                    ? Convert.ToDecimal(dict["BillingProcRcvdQty"])
                    : (dict.ContainsKey("RcvdQty") && dict["RcvdQty"] != null ? Convert.ToDecimal(dict["RcvdQty"]) : 0);

                decimal unAuthValue = dict.ContainsKey("RcvdValueUnAuth") && dict["RcvdValueUnAuth"] != null ? Convert.ToDecimal(dict["RcvdValueUnAuth"]) : 0;
                decimal percentRcvd = dict.ContainsKey("PercentRcvd") && dict["PercentRcvd"] != null ? Convert.ToDecimal(dict["PercentRcvd"]) : 100;
                string lotNo = dict.ContainsKey("LotNo") && dict["LotNo"] != null ? dict["LotNo"].ToString()! : string.Empty;
                string process = dict.ContainsKey("Description") && dict["Description"] != null ? dict["Description"].ToString()! : string.Empty;
                string custCode = dict.ContainsKey("CustCode") && dict["CustCode"] != null ? dict["CustCode"].ToString()! : string.Empty;
                bool audited = dict.ContainsKey("Audited") && dict["Audited"] != null && Convert.ToBoolean(dict["Audited"]);

                decimal balValue = Math.Round(rcvdValue - unAuthValue - repairAmt, 2);

                items.Add(new MakerBillingLotItem
                {
                    EntryID = rcvDetailEntryId,
                    VI_EntryID = viEntryId,
                    IssDetail_RefID = issDetailRefId,
                    RcvReceiptID = rcvReceiptId,
                    DT = dt,
                    OrderNo = orderNo,
                    ItemCode = itemCode,
                    ItemName = itemName,
                    IssQty = issQty,
                    Rate = rate,
                    TempAdvance = tempAdvance,
                    RepairAmt = repairAmt,
                    RcvdValue = rcvdValue,
                    RcvdQty = rcvdQty,
                    UnAuthValue = unAuthValue,
                    BalValue = balValue,
                    PercentRcvd = percentRcvd,
                    Audited = audited,
                    LotNo = lotNo,
                    Process = process,
                    CustCode = custCode,
                    IsSelected = false
                });
            }

            return items;
        }

        public async Task<decimal> GetShortTermDeductionAsync(int vendId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT ISNULL(SUM(Amount - ISNULL(AmountCleared, 0)), 0)
                FROM VMakerAdvancesShortDeduction
                INNER JOIN Makers ON VMakerAdvancesShortDeduction.AccNo = Makers.AccNo
                WHERE Makers.VendID = @VendID";
            return await db.ExecuteScalarAsync<decimal>(sql, new { VendID = vendId });
        }

        public async Task<decimal> GetLongTermBalanceAsync(int vendId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT ISNULL(SUM(TotalAmountTaken - ISNULL(TotalAmountCleared, 0)), 0)
                FROM VMakerAdvancesDeduction
                INNER JOIN Makers ON VMakerAdvancesDeduction.AccNo = Makers.AccNo
                WHERE Makers.VendID = @VendID";
            return await db.ExecuteScalarAsync<decimal>(sql, new { VendID = vendId });
        }

        public async Task<List<MakerRepairItem>> GetRepairWastageItemsAsync(MakerBillingFilter filter)
        {
            if (filter.VendID <= 0) return new List<MakerRepairItem>();

            using IDbConnection db = new SqlConnection(ConnectionString);

            string sql1 = @"
                SELECT EntryID, RecieptID, DT, ItemCode, Qty, Rate,
                       0 AS RepairAmt, (Qty * Rate) AS WastageAmt, LotNo, Description AS Process
                FROM VMakerRepair
                WHERE VendID = @VendID AND VchrNo IS NULL
                  AND (ShowInBilling = 0 OR (ShowInBilling = 1 AND DeductFromBilling = 1))
                  AND DT >= @DtFrom AND DT <= @DtTo";

            string sql2 = @"
                SELECT EntryID, RecieptID, DT, ItemCode, Wastage AS Qty, IssRate AS Rate,
                       0 AS RepairAmt, (Wastage * IssRate) AS WastageAmt, LotNo, '' AS Process
                FROM VMakerBilling_WastageToDeduct
                WHERE VendID = @VendID
                  AND DT >= @DtFrom AND DT <= @DtTo";

            var p = new DynamicParameters();
            p.Add("@VendID", filter.VendID);
            p.Add("@DtFrom", filter.DtFrom.Date);
            p.Add("@DtTo", filter.DtTo.Date.AddDays(1).AddTicks(-1));

            var list1 = (await db.QueryAsync<MakerRepairItem>(sql1, p)).ToList();
            var list2 = (await db.QueryAsync<MakerRepairItem>(sql2, p)).ToList();

            var combined = list1.Concat(list2).ToList();
            foreach (var item in combined)
            {
                item.IsSelected = true;
            }
            return combined;
        }

        public async Task<List<MakerMaterialDedItem>> GetMaterialDeductionsAsync(int vendId, string masterPoNo)
        {
            if (vendId <= 0) return new List<MakerMaterialDedItem>();

            using IDbConnection db = new SqlConnection(ConnectionString);

            string sql = @"
                SELECT EntryID, MasterEntryID, IssNo, DT, RMID1, RMName, ItemID AS ItemCode, Price, Qty
                FROM VRawMaterialIssuance
                WHERE VendID = @VendID
                  AND MasterEntryID NOT IN (SELECT RMI_RefID FROM MakerPostedBills_MaterialDeductions WHERE RMI_RefID IS NOT NULL)
                  AND EntryID NOT IN (SELECT RMID_RefID FROM MakerPostedBills_MaterialDeductions WHERE RMID_RefID IS NOT NULL)
                  AND EntryID NOT IN (SELECT RMID_RefID FROM RawMaterialIssuanceDetail_Hidden_From_Maker_Billing)";

            var p = new DynamicParameters();
            p.Add("@VendID", vendId);

            if (!string.IsNullOrWhiteSpace(masterPoNo))
            {
                sql += " AND MasterPONo = @MasterPONo";
                p.Add("@MasterPONo", masterPoNo.Trim());
            }

            sql += " ORDER BY EntryID";

            var list = (await db.QueryAsync<MakerMaterialDedItem>(sql, p)).ToList();
            foreach (var item in list)
            {
                item.IsSelected = true;
            }
            return list;
        }

        public async Task<List<MakerHiddenLotItem>> GetHiddenLotsAsync(int vendId)
        {
            if (vendId <= 0) return new List<MakerHiddenLotItem>();

            using IDbConnection db = new SqlConnection(ConnectionString);

            const string sql = @"
                SELECT EntryID,
                       VRD_RefID,
                       LotNo,
                       RcvdQty AS Qty,
                       ItemCode,
                       ItemName,
                       OrderNo,
                       CustCode,
                       Description AS Process
                FROM VMakerBilling_Receivings_Hidden
                WHERE VendID = @VendID
                ORDER BY OnlyDT DESC";

            return (await db.QueryAsync<MakerHiddenLotItem>(sql, new { VendID = vendId })).ToList();
        }

        public async Task<List<MakerHiddenRMItem>> GetHiddenRMItemsAsync(int vendId)
        {
            if (vendId <= 0) return new List<MakerHiddenRMItem>();

            using IDbConnection db = new SqlConnection(ConnectionString);

            const string sql = @"
                SELECT EntryID, MasterEntryID, IssNo, DT, RMID1, RMName, ItemID AS ItemCode, Price, Qty
                FROM VRawMaterialIssuance
                WHERE VendID = @VendID
                  AND EntryID IN (SELECT RMID_RefID FROM RawMaterialIssuanceDetail_Hidden_From_Maker_Billing)
                ORDER BY EntryID";

            return (await db.QueryAsync<MakerHiddenRMItem>(sql, new { VendID = vendId })).ToList();
        }

        public async Task UpdateLotRateAsync(long refId, decimal rate)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "UPDATE VendIssdDetail SET Rate = @Rate WHERE RefID = @RefID";
            await db.ExecuteAsync(sql, new { Rate = rate, RefID = refId });
        }

        public async Task UpdateMaterialPriceAsync(long entryId, decimal price)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "UPDATE RawMaterialIssuanceDetail SET Price = @Price WHERE EntryID = @EntryID";
            await db.ExecuteAsync(sql, new { Price = price, EntryID = entryId });
        }

        public async Task MarkLotsAsAuditedAsync(List<long> vrdEntryIds, bool audited)
        {
            if (!vrdEntryIds.Any()) return;

            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "UPDATE VendRcvdDetail SET Audited = @Audited WHERE EntryID IN @IDs";
            await db.ExecuteAsync(sql, new { Audited = audited ? 1 : 0, IDs = vrdEntryIds });
        }

        public async Task HideLotFromBillingAsync(long vrdEntryId, string userName, string machineName)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                IF NOT EXISTS (SELECT 1 FROM MakerBilling_Receivings_Hidden WHERE VRD_RefID = @VRD_RefID)
                BEGIN
                    INSERT INTO MakerBilling_Receivings_Hidden (VRD_RefID, UserName, MachineName, DTEntry)
                    VALUES (@VRD_RefID, @UserName, @MachineName, GETDATE())
                END";
            await db.ExecuteAsync(sql, new { VRD_RefID = vrdEntryId, UserName = userName, MachineName = machineName });
        }

        public async Task RestoreHiddenLotAsync(long hiddenEntryId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "DELETE FROM MakerBilling_Receivings_Hidden WHERE EntryID = @EntryID OR VRD_RefID = @EntryID";
            await db.ExecuteAsync(sql, new { EntryID = hiddenEntryId });
        }

        public async Task HideRMFromBillingAsync(long rmidRefId, string userName, string machineName)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                IF NOT EXISTS (SELECT 1 FROM RawMaterialIssuanceDetail_Hidden_From_Maker_Billing WHERE RMID_RefID = @RMID_RefID)
                BEGIN
                    INSERT INTO RawMaterialIssuanceDetail_Hidden_From_Maker_Billing (RMID_RefID, UserName, MachineName, DTEntry)
                    VALUES (@RMID_RefID, @UserName, @MachineName, GETDATE())
                END";
            await db.ExecuteAsync(sql, new { RMID_RefID = rmidRefId, UserName = userName, MachineName = machineName });
        }

        public async Task RestoreHiddenRMAsync(long rmidRefId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "DELETE FROM RawMaterialIssuanceDetail_Hidden_From_Maker_Billing WHERE RMID_RefID = @RMID_RefID";
            await db.ExecuteAsync(sql, new { RMID_RefID = rmidRefId });
        }

        public async Task<string?> GetItemPicBase64Async(string itemCode)
        {
            if (string.IsNullOrWhiteSpace(itemCode)) return null;

            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT ItemPic FROM Items WHERE ItemID = @ItemID";
            byte[]? bytes = await db.QueryFirstOrDefaultAsync<byte[]>(sql, new { ItemID = itemCode.Trim() });
            if (bytes != null && bytes.Length > 0)
            {
                return $"data:image/jpeg;base64,{Convert.ToBase64String(bytes)}";
            }
            return null;
        }

        public async Task<PostMakerBillResult> SaveMakerPostedBillTransactionAsync(PostMakerBillRequest request, string voucherNo)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var tx = db.BeginTransaction();
            try
            {
                long nextBillNo = await db.ExecuteScalarAsync<long>(@"
                    SELECT ISNULL(MAX(BillNo), 0) + 1
                    FROM MakerPostedBills
                    WHERE BillNo IS NOT NULL AND VendID = @VendID",
                    new { VendID = request.VendID }, tx);

                decimal totalValue = request.SelectedLots.Sum(x => x.Value);
                decimal repairRelax = request.RejectionRelaxation ? (totalValue / 100m) * 2m : 0;
                decimal repairAmtToDeduct = Math.Max(0, request.RepairAmount - repairRelax);

                // 1. Insert MakerPostedBills Header
                const string billSql = @"
                    INSERT INTO MakerPostedBills (
                        VendID, DT, VchrNo, UserName, MachineName, BillNo,
                        TempAdvance, STDeduction, LTDeduction, MaterialDedAmt, ManualDedAmt,
                        DTFrom, DTTo, WastageAmtRelaxation, TotalWastageAmt, BalanceOrderValue, UnAuthValue
                    ) VALUES (
                        @VendID, @DT, @VchrNo, @UserName, @MachineName, @BillNo,
                        0, @STDeduction, @LTDeduction, @MaterialDedAmt, @ManualDedAmt,
                        @DTFrom, @DTTo, @WastageAmtRelaxation, @TotalWastageAmt, 0, 0
                    )";

                await db.ExecuteAsync(billSql, new
                {
                    VendID = request.VendID,
                    DT = request.PostingDate,
                    VchrNo = voucherNo,
                    UserName = request.UserName,
                    MachineName = request.MachineName,
                    BillNo = (int)nextBillNo,
                    STDeduction = (int)request.ShortTermDeduction,
                    LTDeduction = (int)request.LongTermDeduction,
                    MaterialDedAmt = (int)request.MaterialDeductionAmount,
                    ManualDedAmt = (int)request.ManualDeductionAmount,
                    DTFrom = request.DtFrom.Date,
                    DTTo = request.DtTo.Date,
                    WastageAmtRelaxation = (int)repairRelax,
                    TotalWastageAmt = (int)request.RepairAmount
                }, tx);

                long billEntryId = await db.ExecuteScalarAsync<long>("SELECT MAX(EntryID) FROM MakerPostedBills", transaction: tx);

                // 2. Insert Details & Receivings mapping
                foreach (var lot in request.SelectedLots)
                {
                    const string detailSql = @"
                        INSERT INTO MakerPostedBillsDetail (
                            RefID, Iss_RefID, TempAdvance, RepairAmt, RcvdValue,
                            UnAuthValue, BalanceValue, RcvdPercentage, RcvdQty, IssDetail_RefID
                        ) VALUES (
                            @RefID, @Iss_RefID, 0, 0, @RcvdValue,
                            0, 0, 0, @RcvdQty, @IssDetail_RefID
                        )";

                    await db.ExecuteAsync(detailSql, new
                    {
                        RefID = billEntryId,
                        Iss_RefID = lot.VI_EntryID,
                        RcvdValue = (int)lot.Value,
                        RcvdQty = (int)lot.Qty,
                        IssDetail_RefID = lot.IssDetail_RefID
                    }, tx);

                    const string rcvSql = @"
                        INSERT INTO MakerPostedBillsDetail_Receivings (
                            MPB_RefID, VRD_RefID, Qty, Rate, Main_VI_EntryID
                        ) VALUES (
                            @MPB_RefID, @VRD_RefID, @Qty, @Rate, @Main_VI_EntryID
                        )";

                    await db.ExecuteAsync(rcvSql, new
                    {
                        MPB_RefID = billEntryId,
                        VRD_RefID = lot.EntryID,
                        Qty = (int)lot.Qty,
                        Rate = (float)lot.Rate,
                        Main_VI_EntryID = lot.VI_EntryID
                    }, tx);
                }

                // 3. Insert Material Deductions
                if (request.SelectedMaterialDeductions != null)
                {
                    foreach (var mat in request.SelectedMaterialDeductions)
                    {
                        const string matSql = @"
                            INSERT INTO MakerPostedBills_MaterialDeductions (MPB_RefID, RMI_RefID, RMID_RefID)
                            VALUES (@MPB_RefID, @RMI_RefID, @RMID_RefID)";

                        await db.ExecuteAsync(matSql, new
                        {
                            MPB_RefID = billEntryId,
                            RMI_RefID = mat.MasterEntryID,
                            RMID_RefID = mat.EntryID
                        }, tx);
                    }
                }

                // 4. Insert Manual Deductions
                if (request.ManualDeductions != null)
                {
                    foreach (var man in request.ManualDeductions)
                    {
                        const string manSql = @"
                            INSERT INTO MakerPostedBills_ManualDeductions (MPB_RefID, Description, Amount)
                            VALUES (@MPB_RefID, @Description, @Amount)";

                        await db.ExecuteAsync(manSql, new
                        {
                            MPB_RefID = billEntryId,
                            Description = man.Description,
                            Amount = (int)man.Amount
                        }, tx);
                    }
                }

                tx.Commit();

                return new PostMakerBillResult
                {
                    Success = true,
                    BillEntryID = billEntryId,
                    BillNo = nextBillNo,
                    VoucherNo = voucherNo,
                    Message = $"Maker Bill #{nextBillNo} posted successfully with JV #{voucherNo}."
                };
            }
            catch (Exception ex)
            {
                tx.Rollback();
                return new PostMakerBillResult
                {
                    Success = false,
                    Message = ex.Message
                };
            }
        }
    }
}
