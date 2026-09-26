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

        public async Task<List<string>> GetHubNamesAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT Hub_Name FROM Hub_Names ORDER BY Hub_Name";
            return (await db.QueryAsync<string>(sql)).ToList();
        }

        public async Task<bool> AddHubNameAsync(string hubName)
        {
            if (string.IsNullOrWhiteSpace(hubName)) return false;
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                IF NOT EXISTS (SELECT 1 FROM Hub_Names WHERE Hub_Name = @Hub_Name)
                BEGIN
                    INSERT INTO Hub_Names (Hub_Name) VALUES (@Hub_Name);
                END";
            await db.ExecuteAsync(sql, new { Hub_Name = hubName.Trim() });
            return true;
        }

        public async Task<(bool CanDelete, string Reason)> CanDeleteHubNameAsync(string hubName)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string checkSql = "SELECT COUNT(*) FROM ProcessGroupsProcesses WHERE Hub_Name = @Hub_Name";
            int count = await db.ExecuteScalarAsync<int>(checkSql, new { Hub_Name = hubName.Trim() });
            if (count > 0)
            {
                return (false, $"Cannot delete Hub '{hubName}' because it is currently assigned to {count} process sequence step(s).");
            }
            return (true, string.Empty);
        }

        public async Task<bool> DeleteHubNameAsync(string hubName)
        {
            var (canDelete, _) = await CanDeleteHubNameAsync(hubName);
            if (!canDelete) return false;

            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "DELETE FROM Hub_Names WHERE Hub_Name = @Hub_Name";
            int affected = await db.ExecuteAsync(sql, new { Hub_Name = hubName.Trim() });
            return affected > 0;
        }

        public async Task<List<ProcessGroupHubOverviewDto>> GetGroupHubOverviewAsync(int groupId)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);

            // 1. Get distinct hubs with processes in this group
            const string hubProcessSql = @"
                SELECT 
                    pgp.Hub_Name,
                    ISNULL(p.Description, '') AS ProcessName,
                    pgp.SeqNo
                FROM ProcessGroupsProcesses pgp
                LEFT JOIN Processes p ON pgp.Process_RefID = p.ProcessID
                WHERE pgp.Group_RefID = @GroupId AND ISNULL(pgp.Hub_Name, '') <> ''
                ORDER BY pgp.SeqNo";

            var rows = (await db.QueryAsync(hubProcessSql, new { GroupId = groupId })).ToList();

            // 2. Get existing assigned supervisors
            const string supervisorSql = @"
                SELECT 
                    s.GroupID,
                    s.Hub_Name,
                    s.UserID,
                    s.UserName,
                    e.EmpID,
                    e.Name AS EmployeeName,
                    e.Designation
                FROM ProcessGroup_Hub_Supervisors s
                LEFT JOIN Users u ON s.UserID = u.UserID
                LEFT JOIN Employees e ON u.EmpID = e.EmpID
                WHERE s.GroupID = @GroupId
                ORDER BY s.Hub_Name, e.Name";

            var supervisors = (await db.QueryAsync<HubSupervisorDto>(supervisorSql, new { GroupId = groupId })).ToList();
            var supervisorsByHub = supervisors.GroupBy(s => s.Hub_Name, StringComparer.OrdinalIgnoreCase)
                                              .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

            var hubGroups = rows.GroupBy(r => (string)r.Hub_Name, StringComparer.OrdinalIgnoreCase);
            var result = new List<ProcessGroupHubOverviewDto>();

            foreach (var hg in hubGroups)
            {
                var hubOverview = new ProcessGroupHubOverviewDto
                {
                    Hub_Name = hg.Key,
                    ProcessNames = hg.Select(r => (string)r.ProcessName).Distinct().ToList(),
                    Supervisors = supervisorsByHub.TryGetValue(hg.Key, out var sups) ? sups : new List<HubSupervisorDto>()
                };
                result.Add(hubOverview);
            }

            return result;
        }

        public async Task<bool> SaveGroupHubSupervisorsAsync(int groupId, string hubName, List<int> userIds)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            if (db.State != ConnectionState.Open) db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                const string delSql = "DELETE FROM ProcessGroup_Hub_Supervisors WHERE GroupID = @GroupID AND Hub_Name = @Hub_Name";
                await db.ExecuteAsync(delSql, new { GroupID = groupId, Hub_Name = hubName }, trans);

                if (userIds != null && userIds.Any())
                {
                    const string insertSql = @"
                        INSERT INTO ProcessGroup_Hub_Supervisors (GroupID, Hub_Name, UserID, UserName, AssignedAt)
                        SELECT @GroupID, @Hub_Name, u.UserID, u.UserName, GETDATE()
                        FROM Users u
                        WHERE u.UserID = @UserID";

                    foreach (var uid in userIds.Distinct())
                    {
                        await db.ExecuteAsync(insertSql, new { GroupID = groupId, Hub_Name = hubName, UserID = uid }, trans);
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
