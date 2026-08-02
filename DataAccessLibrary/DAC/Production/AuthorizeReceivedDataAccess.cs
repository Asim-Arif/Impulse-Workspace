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
    public class AuthorizeReceivedDataAccess : IAuthorizeReceivedDataAccess
    {
        private readonly IConfiguration _config;

        public AuthorizeReceivedDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection string is not configured.");

        public async Task<List<PendingAuthorizeReceivedItemModel>> GetPendingAuthorizeReceivedItemsAsync(AuthorizeReceivedFilterModel filter)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT VR.EntryID, VR.RefID, VR.RecieptID, ISNULL(VR.VendID, 0) AS VendID,
                                  ISNULL(VR.VendID1, '') AS VendID1, ISNULL(VR.VenderName, '') AS VenderName,
                                  ISNULL(VR.Supervisor, '') AS Supervisor, ISNULL(VR.OrderNo, '') AS OrderNo,
                                  VR.ItemCode, ISNULL(VR.ItemDescription, '') AS ItemDescription,
                                  ISNULL(VR.LotNo, '') AS LotNo, ISNULL(VR.ProcessID, 0) AS ProcessID,
                                  ISNULL(VR.Process, '') AS Process, ISNULL(VR.RcvdQty, 0) AS RcvdQty,
                                  ISNULL(VR.Wastage, 0) AS Wastage, ISNULL(VR.LostQty, 0) AS LostQty,
                                  ISNULL(VR.ReWorkQty, 0) AS ReWorkQty, ISNULL(VR.RepairAmt, 0) AS RepairAmt,
                                  ISNULL(VR.Rate, 0) AS Rate, ISNULL(VR.Issue_RefID, 0) AS Issue_RefID,
                                  ISNULL(VR.Issuance_RefID, 0) AS Issuance_RefID, ISNULL(VR.RcvdWeight, 0) AS RcvdWeight,
                                  ISNULL(VR.Unit, 'Pcs') AS Unit, (SELECT ISNULL(Insp_EmpID, '') FROM VendRcvdDetail WHERE EntryID = VR.EntryID) AS Insp_EmpID
                           FROM VVendRcvItemsrpt VR
                           WHERE VR.ReqAuth = 1
                             AND (@VendID = 0 OR VR.VendID = @VendID)
                             AND (@ProcessID = 0 OR VR.ProcessID = @ProcessID)
                             AND (@LotNo = '' OR VR.LotNo = @LotNo)
                             AND (@DateFrom IS NULL OR VR.DT >= @DateFrom)
                             AND (@DateTo IS NULL OR VR.DT <= @DateTo)
                             AND (@SearchText = '' OR VR.RecieptID LIKE '%' + @SearchText + '%' OR VR.ItemCode LIKE '%' + @SearchText + '%' OR VR.ItemDescription LIKE '%' + @SearchText + '%')
                           ORDER BY VR.EntryID";

            var items = (await db.QueryAsync<PendingAuthorizeReceivedItemModel>(sql, new
            {
                VendID = filter.VendID,
                ProcessID = filter.ProcessID,
                LotNo = filter.LotNo ?? "",
                DateFrom = filter.DateFrom,
                DateTo = filter.DateTo,
                SearchText = filter.SearchText ?? ""
            })).ToList();

            foreach (var item in items)
            {
                item.IsChecked = true;
            }

            return items;
        }

        public async Task<List<WastageTypeLookupModel>> GetWastageTypesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT EntryID, WastageName AS Description, ISNULL(WastageType, 1) AS WastageType FROM WastageTypes WHERE ISNULL(Closed, 0) = 0 ORDER BY WastageName";
            return (await db.QueryAsync<WastageTypeLookupModel>(sql)).ToList();
        }

        public async Task<List<RepairTypeLookupModel>> GetRepairTypesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT EntryID, RepairType AS Description FROM RepairTypes ORDER BY RepairType";
            return (await db.QueryAsync<RepairTypeLookupModel>(sql)).ToList();
        }

        public async Task<List<EmployeeLookupModel>> GetInspectorsAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT EmpID, Name, ISNULL(Designation, '') AS Designation FROM Employees WHERE ISNULL(Active, 1) = 1 ORDER BY Name";
            return (await db.QueryAsync<EmployeeLookupModel>(sql)).ToList();
        }

        public async Task<List<ProcessPOLookupModel>> GetProcessesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT ProcessID, Description, ISNULL(SNO, 0) AS SNO FROM Processes ORDER BY SNO";
            return (await db.QueryAsync<ProcessPOLookupModel>(sql)).ToList();
        }

        public async Task<List<MakerPOLookupModel>> GetMakersAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            string sql = @"SELECT VendID, ISNULL(VendID1, '') AS VendID1, VenderName FROM VMakers WHERE Active = 1 ORDER BY VenderName";
            return (await db.QueryAsync<MakerPOLookupModel>(sql)).ToList();
        }

        public async Task<bool> SaveAuthorizationAsync(List<PendingAuthorizeReceivedItemModel> items, string inspectorEmpId, string userName, string machineName, int userId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                foreach (var item in items.Where(i => i.IsChecked))
                {
                    decimal totalWastage = item.WastageBreakdowns.Any() ? item.WastageBreakdowns.Sum(w => w.Qty) : item.Wastage;
                    decimal totalReWork = item.ReWorkBreakdowns.Any() ? item.ReWorkBreakdowns.Sum(r => r.Qty) : item.ReWorkQty;

                    // 1. Update VendRcvdDetail
                    string updateVrdSql = @"UPDATE VendRcvdDetail 
                                            SET ReqAuth = 0, 
                                                Wastage = @Wastage, 
                                                RepairAmt = @RepairAmt, 
                                                LostQty = @LostQty, 
                                                ReWorkQty = @ReWorkQty, 
                                                Insp_EmpID = @InspEmpID 
                                            WHERE EntryID = @EntryID";

                    await db.ExecuteAsync(updateVrdSql, new
                    {
                        Wastage = totalWastage,
                        RepairAmt = item.RepairAmt,
                        LostQty = item.LostQty,
                        ReWorkQty = totalReWork,
                        InspEmpID = inspectorEmpId ?? "",
                        EntryID = item.EntryID
                    }, trans);

                    // 2. Insert Rework Details
                    if (item.ReWorkBreakdowns != null && item.ReWorkBreakdowns.Any())
                    {
                        foreach (var rw in item.ReWorkBreakdowns)
                        {
                            string insertRwSql = @"INSERT INTO VendRcvdDetailReWorkDetail (VRD_RefID, Repair_RefID, Qty) 
                                                   VALUES (@VRD_RefID, @Repair_RefID, @Qty);
                                                   SELECT SCOPE_IDENTITY();";

                            long rwEntryId = await db.ExecuteScalarAsync<long>(insertRwSql, new
                            {
                                VRD_RefID = item.EntryID,
                                Repair_RefID = rw.Repair_RefID,
                                Qty = rw.Qty
                            }, trans);

                            string insertRwProcSql = @"INSERT INTO VendRcvdDetailReWorkDetailProcesses (VRDRWD_RefID, ProcessID)
                                                       SELECT @VRDRWD_RefID, ProcessID 
                                                       FROM RepairTypeProcesses 
                                                       WHERE Repair_RefID = @Repair_RefID";

                            await db.ExecuteAsync(insertRwProcSql, new { VRDRWD_RefID = rwEntryId, Repair_RefID = rw.Repair_RefID }, trans);
                        }
                    }

                    // 3. Insert Wastage Details & MakerRepair Ledger
                    decimal wastageCh2Vendor = 0;
                    decimal wastageCh2Originator = 0;
                    decimal wastageCh2Selection = 0;

                    if (item.WastageBreakdowns != null && item.WastageBreakdowns.Any())
                    {
                        foreach (var w in item.WastageBreakdowns)
                        {
                            long targetVendId = item.VendID;
                            if (w.WastageType == 2)
                            {
                                wastageCh2Originator += w.Qty;
                                string origVendSql = @"SELECT ISNULL(VR.VendID, 0) 
                                                       FROM VendReceived VR 
                                                       INNER JOIN VendRcvdDetail VRD ON VR.EntryID = VRD.RefID 
                                                       WHERE VRD.EntryID = (SELECT MIN(EntryID) FROM VendRcvdDetail WHERE LotNo = @LotNo)";
                                targetVendId = await db.ExecuteScalarAsync<long?>(origVendSql, new { LotNo = item.LotNo }, trans) ?? item.VendID;
                            }
                            else if (w.WastageType == 3)
                            {
                                wastageCh2Selection += w.Qty;
                                targetVendId = w.VendID > 0 ? w.VendID : item.VendID;
                            }
                            else
                            {
                                wastageCh2Vendor += w.Qty;
                            }

                            string insertWastageSql = @"INSERT INTO VendRcvdDetailWastageDetail (
                                                            VRD_RefID, Wastage_RefID, Qty, WastageType, ReturnTo_VRD_RefID, VendID, EmpID
                                                        ) VALUES (
                                                            @VRD_RefID, @Wastage_RefID, @Qty, @WastageType, @ReturnTo_VRD_RefID, @VendID, @EmpID
                                                        )";

                            await db.ExecuteAsync(insertWastageSql, new
                            {
                                VRD_RefID = item.EntryID,
                                Wastage_RefID = w.Wastage_RefID,
                                Qty = w.Qty,
                                WastageType = w.WastageType,
                                ReturnTo_VRD_RefID = w.ReturnTo_VRD_RefID,
                                VendID = targetVendId,
                                EmpID = w.EmpID ?? ""
                            }, trans);
                        }
                    }

                    // 4. MakerRepair Ledger Insertions
                    if ((wastageCh2Vendor > 0 || item.RepairAmt > 0) && item.EntryID > 0)
                    {
                        string issueRateSql = @"SELECT ISNULL(Rate, 0) FROM VVendIssdDetail WHERE EntryID = @IssueRefID";
                        double issueRate = await db.ExecuteScalarAsync<double?>(issueRateSql, new { IssueRefID = item.Issue_RefID }, trans) ?? (double)item.Rate;

                        string insertMrSql = @"INSERT INTO MakerRepair (
                                                   VendID, DT, Rcvd_RefID, Qty, Rate, UserName, MachineName, VendIssdDetail_RefID, RepairAmt, VRD_RefID_ChargeTo
                                               ) VALUES (
                                                   @VendID, GETDATE(), @Rcvd_RefID, @Qty, @Rate, @UserName, @MachineName, @VendIssdDetail_RefID, @RepairAmt, @VRD_RefID_ChargeTo
                                               )";

                        await db.ExecuteAsync(insertMrSql, new
                        {
                            VendID = item.VendID,
                            Rcvd_RefID = item.EntryID,
                            Qty = wastageCh2Vendor,
                            Rate = issueRate,
                            UserName = userName,
                            MachineName = machineName,
                            VendIssdDetail_RefID = item.Issue_RefID,
                            RepairAmt = item.RepairAmt,
                            VRD_RefID_ChargeTo = item.EntryID
                        }, trans);
                    }

                    // 5. Insert Audit Log
                    string insertAuthLogSql = @"INSERT INTO VendRcvAuthDetail (UserID, RcvID, DT) VALUES (@UserID, @RcvID, GETDATE())";
                    await db.ExecuteAsync(insertAuthLogSql, new { UserID = userId, RcvID = item.EntryID }, trans);

                    // 6. Update Issued Quantity (Release Lot)
                    if (item.Issue_RefID > 0)
                    {
                        string updateVidSql = @"UPDATE VendIssdDetail 
                                                SET RcvdQty = ISNULL(RcvdQty, 0) + @RcvdQty 
                                                WHERE EntryID = @IssueRefID";

                        await db.ExecuteAsync(updateVidSql, new { RcvdQty = item.RcvdQty, IssueRefID = item.Issue_RefID }, trans);
                    }
                }

                trans.Commit();
                return true;
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }
    }
}
