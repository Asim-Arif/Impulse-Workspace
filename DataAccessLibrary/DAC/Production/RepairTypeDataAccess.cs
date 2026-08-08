using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Production
{
    public class RepairTypeDataAccess : IRepairTypeDataAccess
    {
        private readonly IConfiguration _configuration;

        public RepairTypeDataAccess(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        private string ConnectionString => _configuration.GetConnectionString("DefaultConnection")
            ?? _configuration.GetConnectionString("ImpulseDatabase")
            ?? string.Empty;

        public async Task<List<RepairTypeListItem>> GetRepairTypesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            const string masterSql = @"
                SELECT 
                    EntryID,
                    ISNULL(RepairCode, '') AS RepairCode,
                    ISNULL(RepairType, '') AS RepairType,
                    ISNULL(DefectType, '') AS DefectType,
                    ISNULL(ReturnToOriginator, 0) AS ReturnToOriginator
                FROM RepairTypes
                ORDER BY EntryID";

            var repairTypes = (await db.QueryAsync<RepairTypeListItem>(masterSql)).ToList();

            const string childSql = @"
                SELECT 
                    RTP_EntryID AS EntryID,
                    EntryID AS Repair_RefID,
                    ProcessID,
                    ISNULL(Code, '') AS Code,
                    ISNULL(Description, '') AS Description,
                    ISNULL(SeqNo, 0) AS SeqNo,
                    ISNULL(Scanning, 0) AS Scanning
                FROM VRepairTypeProcesses
                ORDER BY SeqNo, RTP_EntryID";

            var allChildren = (await db.QueryAsync<RepairTypeProcessDto>(childSql)).ToList();
            var groupedChildren = allChildren.GroupBy(c => c.Repair_RefID).ToDictionary(g => g.Key, g => g.ToList());

            foreach (var rt in repairTypes)
            {
                if (groupedChildren.TryGetValue(rt.EntryID, out var children))
                {
                    rt.Processes = children;
                    rt.ProcessCount = children.Count;
                    rt.ProcessSummary = string.Join(" → ", children.Select(c => $"[{c.Code}] {c.Description}"));
                }
            }

            return repairTypes;
        }

        public async Task<RepairTypeDto?> GetRepairTypeByIdAsync(int entryId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            const string masterSql = @"
                SELECT 
                    EntryID,
                    ISNULL(RepairCode, '') AS RepairCode,
                    ISNULL(RepairType, '') AS RepairType,
                    ISNULL(DefectType, '') AS DefectType,
                    ISNULL(ReturnToOriginator, 0) AS ReturnToOriginator
                FROM RepairTypes
                WHERE EntryID = @EntryID";

            var dto = await db.QueryFirstOrDefaultAsync<RepairTypeDto>(masterSql, new { EntryID = entryId });
            if (dto == null) return null;

            dto.Processes = await GetAvailableProcessesForRepairAsync(entryId);
            return dto;
        }

        public async Task<List<RepairTypeProcessItemDto>> GetAvailableProcessesForRepairAsync(int? repairEntryId = null)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            const string allProcessesSql = @"
                SELECT 
                    ProcessID,
                    ISNULL(Code, '') AS Code,
                    ISNULL(Description, '') AS Description
                FROM Processes
                ORDER BY SNo, ProcessID";

            var allProcesses = (await db.QueryAsync<RepairTypeProcessItemDto>(allProcessesSql)).ToList();

            if (repairEntryId.HasValue && repairEntryId.Value > 0)
            {
                const string selectedSql = @"
                    SELECT 
                        ProcessID,
                        ISNULL(SeqNo, 0) AS SeqNo,
                        ISNULL(Scanning, 0) AS Scanning
                    FROM RepairTypeProcesses
                    WHERE Repair_RefID = @Repair_RefID";

                var selectedRows = (await db.QueryAsync<RepairTypeProcessItemDto>(selectedSql, new { Repair_RefID = repairEntryId.Value })).ToList();
                var selectedDict = selectedRows.ToDictionary(r => r.ProcessID, r => r);

                foreach (var p in allProcesses)
                {
                    if (selectedDict.TryGetValue(p.ProcessID, out var match))
                    {
                        p.IsSelected = true;
                        p.SeqNo = match.SeqNo;
                        p.Scanning = match.Scanning;
                    }
                }
            }

            return allProcesses;
        }

        public async Task<int> SaveRepairTypeAsync(RepairTypeDto model)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            if (db.State != ConnectionState.Open) db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                const string insertMasterSql = @"
                    INSERT INTO RepairTypes (
                        RepairCode, RepairType, DefectType, ReturnToOriginator
                    ) VALUES (
                        @RepairCode, @RepairType, @DefectType, @ReturnToOriginator
                    );
                    SELECT CAST(SCOPE_IDENTITY() AS INT);";

                int repairId = await db.ExecuteScalarAsync<int>(insertMasterSql, new
                {
                    RepairCode = model.RepairCode ?? string.Empty,
                    RepairType = model.RepairType ?? string.Empty,
                    DefectType = model.DefectType ?? string.Empty,
                    ReturnToOriginator = model.ReturnToOriginator ? 1 : 0
                }, trans);

                int autoSeq = 1;
                var selectedProcesses = model.Processes.Where(p => p.IsSelected).OrderBy(p => p.SeqNo > 0 ? p.SeqNo : autoSeq).ToList();

                foreach (var p in selectedProcesses)
                {
                    const string insertChildSql = @"
                        INSERT INTO RepairTypeProcesses (
                            Repair_RefID, ProcessID, SeqNo, Scanning
                        ) VALUES (
                            @Repair_RefID, @ProcessID, @SeqNo, @Scanning
                        );";

                    await db.ExecuteAsync(insertChildSql, new
                    {
                        Repair_RefID = repairId,
                        ProcessID = p.ProcessID,
                        SeqNo = p.SeqNo > 0 ? p.SeqNo : autoSeq++,
                        Scanning = p.Scanning ? 1 : 0
                    }, trans);
                }

                trans.Commit();
                return repairId;
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }

        public async Task<bool> UpdateRepairTypeAsync(RepairTypeDto model)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            if (db.State != ConnectionState.Open) db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                const string updateMasterSql = @"
                    UPDATE RepairTypes SET 
                        RepairCode = @RepairCode,
                        RepairType = @RepairType,
                        DefectType = @DefectType,
                        ReturnToOriginator = @ReturnToOriginator
                    WHERE EntryID = @EntryID;";

                await db.ExecuteAsync(updateMasterSql, new
                {
                    RepairCode = model.RepairCode ?? string.Empty,
                    RepairType = model.RepairType ?? string.Empty,
                    DefectType = model.DefectType ?? string.Empty,
                    ReturnToOriginator = model.ReturnToOriginator ? 1 : 0,
                    EntryID = model.EntryID
                }, trans);

                const string deleteChildrenSql = "DELETE FROM RepairTypeProcesses WHERE Repair_RefID = @EntryID;";
                await db.ExecuteAsync(deleteChildrenSql, new { EntryID = model.EntryID }, trans);

                int autoSeq = 1;
                var selectedProcesses = model.Processes.Where(p => p.IsSelected).OrderBy(p => p.SeqNo > 0 ? p.SeqNo : autoSeq).ToList();

                foreach (var p in selectedProcesses)
                {
                    const string insertChildSql = @"
                        INSERT INTO RepairTypeProcesses (
                            Repair_RefID, ProcessID, SeqNo, Scanning
                        ) VALUES (
                            @Repair_RefID, @ProcessID, @SeqNo, @Scanning
                        );";

                    await db.ExecuteAsync(insertChildSql, new
                    {
                        Repair_RefID = model.EntryID,
                        ProcessID = p.ProcessID,
                        SeqNo = p.SeqNo > 0 ? p.SeqNo : autoSeq++,
                        Scanning = p.Scanning ? 1 : 0
                    }, trans);
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

        public async Task<(bool CanDelete, string Reason)> CanDeleteRepairTypeAsync(int entryId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            int usageCount = await db.ExecuteScalarAsync<int>(
                "SELECT COUNT(1) FROM VendRcvdDetailReWorkDetail WHERE Repair_RefID = @EntryID",
                new { EntryID = entryId });

            if (usageCount > 0)
                return (false, $"Cannot delete repair type: It is linked to {usageCount} lot rework receiving record(s).");

            return (true, string.Empty);
        }

        public async Task<bool> DeleteRepairTypeAsync(int entryId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            if (db.State != ConnectionState.Open) db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                await db.ExecuteAsync("DELETE FROM RepairTypeProcesses WHERE Repair_RefID = @EntryID;", new { EntryID = entryId }, trans);
                int rows = await db.ExecuteAsync("DELETE FROM RepairTypes WHERE EntryID = @EntryID;", new { EntryID = entryId }, trans);
                trans.Commit();
                return rows > 0;
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }
    }
}
