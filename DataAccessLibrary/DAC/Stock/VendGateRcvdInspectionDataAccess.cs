using Dapper;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Stock
{
    public class VendGateRcvdInspectionDataAccess : IVendGateRcvdInspectionDataAccess
    {
        private readonly string _connectionString;

        public VendGateRcvdInspectionDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<List<VendGateRcvdInspectionMaterialItem>> GetMaterialsByRcvIDAsync(string rcvId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT 
                        d.EntryID,
                        '{' + v.RMID1 + '} ' + v.RMName AS MaterialName,
                        v.RMSize,
                        v.GroupID,
                        ISNULL(d.QtyRcvd, 0) AS QtyRcvd,
                        hdr.PORefNo AS OrderNo,
                        hdr.AccTitle AS VendorName,
                        hdr.DT AS RcvDate
                    FROM VendRcvdDetailPO d
                    INNER JOIN VRMAndItemsUnioned v ON d.MaterialID = v.RMID1
                    INNER JOIN VVendRcvdDetailPO hdr ON d.RcvID = hdr.RcvID
                    WHERE d.RcvID = @RcvID";

                var results = await db.QueryAsync<VendGateRcvdInspectionMaterialItem>(sql, new { RcvID = rcvId });
                return results.ToList();
            }
        }

        public async Task<VendGateRcvdInspectionViewModel> GetInspectionByMaterialAsync(int entryId, int groupId)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var model = new VendGateRcvdInspectionViewModel { VRDPO_RefID = entryId };

                // Check if an inspection already exists
                string sqlHeader = "SELECT * FROM VendRcvdDetailPO_Inspection WHERE VRDPO_RefID = @VRDPO_RefID";
                var existingHeader = await db.QueryFirstOrDefaultAsync<dynamic>(sqlHeader, new { VRDPO_RefID = entryId });

                if (existingHeader != null)
                {
                    model.EntryID = existingHeader.EntryID;
                    model.Disposation = existingHeader.Disposation;
                    model.Comments = existingHeader.Comments;
                    model.DT = existingHeader.DT ?? DateTime.Today;
                    
                    bool status = existingHeader.Status ?? true;
                    // In Legacy: Status = 1 is Pass, Status = 0 is Fail.
                    model.IsPassed = status; 

                    string sqlDetails = @"
                        SELECT 
                            RMGP_RefID,
                            '' AS ParameterName, -- Will update if needed, but we can join RMGroupParameters
                            AQL,
                            SampleSize,
                            RejectOn,
                            TestSpecificationNo,
                            ActualRejection,
                            Status
                        FROM VendRcvdDetailPO_InspectionDetail 
                        WHERE VRDPO_Insp_RefID = @InspRefID";
                    
                    var details = await db.QueryAsync<VendGateRcvdInspectionDetailViewModel>(sqlDetails, new { InspRefID = model.EntryID });
                    
                    // Join with RMGroupParameters to get ParameterName
                    string sqlParams = "SELECT EntryID, ParameterName FROM RMGroupParameters WHERE GroupID = @GroupID";
                    var rmParams = (await db.QueryAsync<dynamic>(sqlParams, new { GroupID = groupId })).ToDictionary(k => (int)k.EntryID, v => (string)v.ParameterName);
                    
                    foreach(var d in details)
                    {
                        if (rmParams.TryGetValue(d.RMGP_RefID, out var pName))
                        {
                            d.ParameterName = pName;
                        }
                    }

                    model.Details = details.ToList();
                }
                else
                {
                    // Generate new details from RMGroupParameters
                    string sqlParams = "SELECT * FROM RMGroupParameters WHERE GroupID = @GroupID ORDER BY EntryID";
                    var rmParams = await db.QueryAsync<dynamic>(sqlParams, new { GroupID = groupId });

                    foreach (var p in rmParams)
                    {
                        model.Details.Add(new VendGateRcvdInspectionDetailViewModel
                        {
                            RMGP_RefID = p.EntryID,
                            ParameterName = p.ParameterName,
                            AQL = p.AQL,
                            SampleSize = p.SampleSize,
                            RejectOn = p.RejectOn,
                            TestSpecificationNo = p.TechSpecNo,
                            ActualRejection = "0",
                            Status = "OK"
                        });
                    }
                }

                return model;
            }
        }

        public async Task SaveInspectionAsync(VendGateRcvdInspectionViewModel model, string userName, string machineName)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                db.Open();
                using (var transaction = db.BeginTransaction())
                {
                    try
                    {
                        // Delete existing if any
                        if (model.EntryID > 0)
                        {
                            await db.ExecuteAsync("DELETE FROM VendRcvdDetailPO_InspectionDetail WHERE VRDPO_Insp_RefID = @EntryID", new { model.EntryID }, transaction);
                            await db.ExecuteAsync("DELETE FROM VendRcvdDetailPO_Inspection WHERE EntryID = @EntryID", new { model.EntryID }, transaction);
                        }
                        else
                        {
                            // In case it wasn't loaded but exists
                            await db.ExecuteAsync(@"
                                DELETE d FROM VendRcvdDetailPO_InspectionDetail d INNER JOIN VendRcvdDetailPO_Inspection i ON d.VRDPO_Insp_RefID = i.EntryID WHERE i.VRDPO_RefID = @VRDPO_RefID;
                                DELETE FROM VendRcvdDetailPO_Inspection WHERE VRDPO_RefID = @VRDPO_RefID;
                            ", new { model.VRDPO_RefID }, transaction);
                        }

                        // Insert new Inspection
                        string insertHeader = @"
                            INSERT INTO VendRcvdDetailPO_Inspection (VRDPO_RefID, Disposation, Comments, DT, Status, UserName, MachineName)
                            VALUES (@VRDPO_RefID, @Disposation, @Comments, @DT, @Status, @UserName, @MachineName);
                            SELECT CAST(SCOPE_IDENTITY() as int);
                        ";

                        bool statusDb = model.IsPassed; // Pass = 1, Fail = 0

                        int newEntryId = await db.QuerySingleAsync<int>(insertHeader, new
                        {
                            model.VRDPO_RefID,
                            model.Disposation,
                            model.Comments,
                            model.DT,
                            Status = statusDb,
                            UserName = userName,
                            MachineName = machineName
                        }, transaction);

                        // Insert Details
                        string insertDetail = @"
                            INSERT INTO VendRcvdDetailPO_InspectionDetail 
                            (VRDPO_Insp_RefID, RMGP_RefID, AQL, SampleSize, RejectOn, TestSpecificationNo, ActualRejection, Status)
                            VALUES 
                            (@VRDPO_Insp_RefID, @RMGP_RefID, @AQL, @SampleSize, @RejectOn, @TestSpecificationNo, @ActualRejection, @Status)
                        ";

                        foreach (var d in model.Details)
                        {
                            await db.ExecuteAsync(insertDetail, new
                            {
                                VRDPO_Insp_RefID = newEntryId,
                                d.RMGP_RefID,
                                d.AQL,
                                d.SampleSize,
                                d.RejectOn,
                                d.TestSpecificationNo,
                                d.ActualRejection,
                                d.Status
                            }, transaction);
                        }

                        transaction.Commit();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }
    }
}
