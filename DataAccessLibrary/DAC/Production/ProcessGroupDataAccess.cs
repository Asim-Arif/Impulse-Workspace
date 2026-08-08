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
    public class ProcessGroupDataAccess : IProcessGroupDataAccess
    {
        private readonly IConfiguration _configuration;

        public ProcessGroupDataAccess(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        private string ConnectionString => _configuration.GetConnectionString("DefaultConnection")
            ?? _configuration.GetConnectionString("ImpulseDatabase")
            ?? string.Empty;

        public async Task<List<ProcessGroupListItem>> GetProcessGroupsAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            const string masterSql = "SELECT EntryID, ISNULL(GroupName, '') AS GroupName FROM ProcessGroups ORDER BY EntryID";
            var groups = (await db.QueryAsync<ProcessGroupListItem>(masterSql)).ToList();

            const string childSql = @"
                SELECT 
                    ChildEntryID AS EntryID,
                    Group_RefID,
                    Process_RefID,
                    ISNULL(Code, '') AS ProcessCode,
                    ISNULL(Description, '') AS ProcessName,
                    ISNULL(Scanning, 0) AS Scanning,
                    ISNULL(Hub_Name, '') AS Hub_Name,
                    ISNULL(SeqNo, 0) AS SeqNo
                FROM VProcessGroups
                ORDER BY Group_RefID, SeqNo, ChildEntryID";

            var allChildren = (await db.QueryAsync<ProcessGroupChildDto>(childSql)).ToList();
            var groupedChildren = allChildren.GroupBy(c => c.Group_RefID).ToDictionary(g => g.Key, g => g.ToList());

            foreach (var g in groups)
            {
                if (groupedChildren.TryGetValue(g.EntryID, out var children))
                {
                    g.Children = children;
                    g.ProcessCount = children.Count;
                    g.FlowSummary = string.Join(" → ", children.Select(c => $"[{c.ProcessCode}] {c.ProcessName}"));
                }
            }

            return groups;
        }

        public async Task<ProcessGroupDto?> GetProcessGroupByIdAsync(int entryId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            const string groupSql = "SELECT EntryID, ISNULL(GroupName, '') AS GroupName FROM ProcessGroups WHERE EntryID = @EntryID";
            var group = await db.QueryFirstOrDefaultAsync<ProcessGroupDto>(groupSql, new { EntryID = entryId });
            if (group == null) return null;

            const string childSql = @"
                SELECT 
                    ChildEntryID AS EntryID,
                    Group_RefID,
                    Process_RefID,
                    ISNULL(Code, '') AS ProcessCode,
                    ISNULL(Description, '') AS ProcessName,
                    ISNULL(Scanning, 0) AS Scanning,
                    ISNULL(Hub_Name, '') AS Hub_Name,
                    ISNULL(SeqNo, 0) AS SeqNo
                FROM VProcessGroups
                WHERE Group_RefID = @EntryID
                ORDER BY SeqNo, ChildEntryID";

            group.Processes = (await db.QueryAsync<ProcessGroupChildDto>(childSql, new { EntryID = entryId })).ToList();
            return group;
        }

        public async Task<int> SaveProcessGroupAsync(ProcessGroupDto model)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            if (db.State != ConnectionState.Open) db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                const string insertMasterSql = @"
                    INSERT INTO ProcessGroups (GroupName) VALUES (@GroupName);
                    SELECT CAST(SCOPE_IDENTITY() AS INT);";

                int groupId = await db.ExecuteScalarAsync<int>(insertMasterSql, new { GroupName = model.GroupName ?? string.Empty }, trans);

                int seq = 1;
                foreach (var p in model.Processes)
                {
                    const string insertChildSql = @"
                        INSERT INTO ProcessGroupsProcesses (
                            Group_RefID, Process_RefID, Scanning, Hub_Name, SeqNo
                        ) VALUES (
                            @Group_RefID, @Process_RefID, @Scanning, @Hub_Name, @SeqNo
                        );";

                    await db.ExecuteAsync(insertChildSql, new
                    {
                        Group_RefID = groupId,
                        Process_RefID = p.Process_RefID,
                        Scanning = p.Scanning ? 1 : 0,
                        Hub_Name = p.Hub_Name ?? string.Empty,
                        SeqNo = p.SeqNo > 0 ? p.SeqNo : seq++
                    }, trans);
                }

                trans.Commit();
                return groupId;
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }

        public async Task<bool> UpdateProcessGroupAsync(ProcessGroupDto model)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            if (db.State != ConnectionState.Open) db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                const string updateMasterSql = "UPDATE ProcessGroups SET GroupName = @GroupName WHERE EntryID = @EntryID";
                await db.ExecuteAsync(updateMasterSql, new { GroupName = model.GroupName ?? string.Empty, EntryID = model.EntryID }, trans);

                const string deleteChildrenSql = "DELETE FROM ProcessGroupsProcesses WHERE Group_RefID = @EntryID";
                await db.ExecuteAsync(deleteChildrenSql, new { EntryID = model.EntryID }, trans);

                int seq = 1;
                foreach (var p in model.Processes)
                {
                    const string insertChildSql = @"
                        INSERT INTO ProcessGroupsProcesses (
                            Group_RefID, Process_RefID, Scanning, Hub_Name, SeqNo
                        ) VALUES (
                            @Group_RefID, @Process_RefID, @Scanning, @Hub_Name, @SeqNo
                        );";

                    await db.ExecuteAsync(insertChildSql, new
                    {
                        Group_RefID = model.EntryID,
                        Process_RefID = p.Process_RefID,
                        Scanning = p.Scanning ? 1 : 0,
                        Hub_Name = p.Hub_Name ?? string.Empty,
                        SeqNo = p.SeqNo > 0 ? p.SeqNo : seq++
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

        public async Task<(bool CanDelete, string Reason)> CanDeleteProcessGroupAsync(int entryId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            int itemCount = await db.ExecuteScalarAsync<int>(
                "SELECT COUNT(1) FROM ItemProcessGroups WHERE PG_RefID = @EntryID",
                new { EntryID = entryId });

            if (itemCount > 0)
                return (false, $"Cannot delete process group: It is currently linked to {itemCount} item(s).");

            return (true, string.Empty);
        }

        public async Task<bool> DeleteProcessGroupAsync(int entryId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            if (db.State != ConnectionState.Open) db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                await db.ExecuteAsync("DELETE FROM ProcessGroupsProcesses WHERE Group_RefID = @EntryID", new { EntryID = entryId }, trans);
                int rows = await db.ExecuteAsync("DELETE FROM ProcessGroups WHERE EntryID = @EntryID", new { EntryID = entryId }, trans);
                trans.Commit();
                return rows > 0;
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }

        public async Task<bool> UpdateItemProcessesFromProcessGroupAsync(int entryId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            int rows = await db.ExecuteAsync(
                "UpdateItemProcessesFromProcessGroups_SP",
                new { PG_RefID = entryId },
                commandType: CommandType.StoredProcedure);

            return true;
        }

        public async Task<List<LookupItemInt>> GetAvailableProcessesLookupAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT ProcessID AS Id, Code + ' - ' + ISNULL(Description, '') AS Name FROM Processes ORDER BY Code";
            var list = (await db.QueryAsync<LookupItemInt>(sql)).ToList();
            return list;
        }
    }
}
