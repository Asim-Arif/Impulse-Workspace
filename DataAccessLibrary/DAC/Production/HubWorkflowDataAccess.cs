using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.Production
{
    public class HubWorkflowDataAccess : IHubWorkflowDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<HubWorkflowDataAccess> _logger;

        public HubWorkflowDataAccess(IConfiguration config, ILogger<HubWorkflowDataAccess> logger)
        {
            _connectionString = config.GetConnectionString("DefaultConnection")
                ?? config.GetConnectionString("SMBI_AWM")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        private SqlConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task<HubDetectionResultDto?> AnalyzeLotProcessHubAsync(string itemCode, int processId)
        {
            try
            {
                using var conn = CreateConnection();

                // 1. Fetch Process Group for this Item
                const string groupSql = @"
                    SELECT TOP 1 ipg.PG_RefID AS GroupID, pg.GroupName
                    FROM ItemProcessGroups ipg
                    INNER JOIN ProcessGroups pg ON ipg.PG_RefID = pg.EntryID
                    WHERE ipg.ItemID = @ItemCode";

                var group = await conn.QueryFirstOrDefaultAsync<dynamic>(groupSql, new { ItemCode = itemCode });
                if (group == null)
                {
                    _logger.LogWarning("No ProcessGroup assigned in ItemProcessGroups for Item [{ItemCode}]", itemCode);
                    return null;
                }

                int groupId = (int)group.GroupID;
                string groupName = (string)group.GroupName;

                // 2. Fetch all processes & hubs for this Process Group
                const string processesSql = @"
                    SELECT 
                        pgp.Process_RefID, 
                        pgp.SeqNo, 
                        ISNULL(pgp.Hub_Name, '') AS Hub_Name, 
                        p.Description AS ProcessName
                    FROM ProcessGroupsProcesses pgp
                    INNER JOIN Processes p ON pgp.Process_RefID = p.ProcessID
                    WHERE pgp.Group_RefID = @GroupID
                    ORDER BY pgp.SeqNo ASC";

                var groupProcesses = (await conn.QueryAsync<dynamic>(processesSql, new { GroupID = groupId })).ToList();
                if (!groupProcesses.Any())
                {
                    _logger.LogWarning("No processes configured in ProcessGroupsProcesses for Group #{GroupID}", groupId);
                    return null;
                }

                // 3. Identify the current process in the group sequence
                var currentProcess = groupProcesses.FirstOrDefault(p => (int)p.Process_RefID == processId);
                if (currentProcess == null)
                {
                    _logger.LogInformation("ProcessID {ProcessID} not found in Group #{GroupID} sequence for Item [{ItemCode}]", processId, groupId, itemCode);
                    return null;
                }

                string currentHubName = (string)currentProcess.Hub_Name;
                int currentSeqNo = (int)currentProcess.SeqNo;
                string currentProcessName = (string)currentProcess.ProcessName;

                // Find the maximum sequence number for this Hub
                var hubProcesses = groupProcesses.Where(p => string.Equals((string)p.Hub_Name, currentHubName, StringComparison.OrdinalIgnoreCase)).ToList();
                int maxSeqInHub = hubProcesses.Max(p => (int)p.SeqNo);

                bool isLastProcessOfHub = (currentSeqNo == maxSeqInHub);

                var result = new HubDetectionResultDto
                {
                    GroupID = groupId,
                    GroupName = groupName,
                    CurrentProcessId = processId,
                    CurrentProcessName = currentProcessName,
                    CurrentSeqNo = currentSeqNo,
                    CurrentHubName = currentHubName,
                    IsLastProcessOfHub = isLastProcessOfHub
                };

                if (!isLastProcessOfHub)
                {
                    // Intermediate process in this hub; no hub transition occurs
                    return result;
                }

                // 4. Find the first process of the next Hub
                var nextProcess = groupProcesses.FirstOrDefault(p => (int)p.SeqNo > currentSeqNo && !string.Equals((string)p.Hub_Name, currentHubName, StringComparison.OrdinalIgnoreCase));

                if (nextProcess != null)
                {
                    result.IsFinalGroupHub = false;
                    result.NextHubName = (string)nextProcess.Hub_Name;
                    result.NextFirstProcessId = (int)nextProcess.Process_RefID;
                    result.NextFirstProcessName = (string)nextProcess.ProcessName;

                    // Fetch supervisors configured for NextHubName
                    const string supervisorsSql = @"
                        SELECT 
                            pghs.UserID, 
                            pghs.UserName, 
                            u.EmpID, 
                            ISNULL(e.Name, pghs.UserName) AS EmployeeName,
                            e.Designation
                        FROM ProcessGroup_Hub_Supervisors pghs
                        INNER JOIN Users u ON pghs.UserID = u.UserID
                        LEFT JOIN Employees e ON u.EmpID = e.EmpID
                        WHERE pghs.GroupID = @GroupID 
                          AND pghs.Hub_Name = @NextHubName
                          AND COALESCE(u.InActive, 0) = 0";

                    var supervisors = (await conn.QueryAsync<HubSupervisorInfoDto>(supervisorsSql, new
                    {
                        GroupID = groupId,
                        NextHubName = result.NextHubName
                    })).ToList();

                    result.NextHubSupervisors = supervisors;
                }
                else
                {
                    result.IsFinalGroupHub = true;
                    result.NextHubName = null;
                }

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error analyzing lot process hub for Item [{ItemCode}], ProcessID {ProcessID}", itemCode, processId);
                return null;
            }
        }

        public async Task<bool> IsProcessAuthRequiredAsync(int processId)
        {
            try
            {
                using var conn = CreateConnection();
                const string sql = "SELECT ISNULL(AuthRequired, 0) FROM Processes WHERE ProcessID = @ProcessId";
                return await conn.ExecuteScalarAsync<bool>(sql, new { ProcessId = processId });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error checking AuthRequired for ProcessID {ProcessId}", processId);
                return false;
            }
        }

        public async Task<int> CloseHubTaskAsync(string lotNo, string hubName, string completedByUserName, string notes)
        {
            try
            {
                using var conn = CreateConnection();
                string entityRefId = $"{lotNo}:{hubName}";

                const string updateSql = @"
                    UPDATE TaskItems
                    SET Status = 2,
                        CompletedAt = GETUTCDATE(),
                        CompletedBy = @CompletedBy,
                        UpdatedAt = GETUTCDATE()
                    WHERE Status <> 2
                      AND (
                          (SourceEntityType = 'LotHub' AND SourceEntityRefId = @EntityRefId)
                          OR (SourceEntityType = 'Lot' AND SourceEntityRefId = @LotNo AND Title LIKE '%' + @HubName + '%')
                      );";

                int rows = await conn.ExecuteAsync(updateSql, new
                {
                    CompletedBy = completedByUserName,
                    EntityRefId = entityRefId,
                    LotNo = lotNo,
                    HubName = hubName
                });

                if (rows > 0)
                {
                    // Mark corresponding notifications read
                    const string notifSql = @"
                        UPDATE AppNotifications
                        SET IsRead = 1
                        WHERE IsRead = 0
                          AND (ActionUrl LIKE '%' + @LotNo + '%' OR Title LIKE '%' + @LotNo + '%' OR Message LIKE '%' + @LotNo + '%');";

                    await conn.ExecuteAsync(notifSql, new
                    {
                        LotNo = lotNo
                    });

                    if (!string.IsNullOrWhiteSpace(notes))
                    {
                        const string commentSql = @"
                            INSERT INTO TaskComments (TaskId, UserId, Content, CreatedAt)
                            SELECT Id, @CompletedBy, @Content, GETUTCDATE()
                            FROM TaskItems
                            WHERE (SourceEntityType = 'LotHub' AND SourceEntityRefId = @EntityRefId)
                               OR (SourceEntityType = 'Lot' AND SourceEntityRefId = @LotNo AND Title LIKE '%' + @HubName + '%');";

                        await conn.ExecuteAsync(commentSql, new
                        {
                            CompletedBy = completedByUserName,
                            Content = notes,
                            EntityRefId = entityRefId,
                            LotNo = lotNo,
                            HubName = hubName
                        });
                    }

                    _logger.LogInformation("Closed {Count} workflow task(s) for Lot [{LotNo}] at Hub [{HubName}]", rows, lotNo, hubName);
                }

                return rows;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error closing hub task for Lot [{LotNo}], Hub [{HubName}]", lotNo, hubName);
                return 0;
            }
        }

        public async Task<int> CreateNextHubTaskAndNotificationsAsync(
            HubDetectionResultDto transition,
            string lotNo,
            string itemCode,
            string orderNo,
            decimal rcvdQty,
            string createdBy)
        {
            try
            {
                using var conn = CreateConnection();
                await conn.OpenAsync();
                using var trans = conn.BeginTransaction();

                string nextHub = transition.NextHubName ?? "Next Station";
                string entityRefId = $"{lotNo}:{nextHub}";
                string title = $"Lot #{lotNo} Ready for Hub {nextHub} (Order #{orderNo})";
                string desc = $"Item [{itemCode}], Lot #{lotNo} ({rcvdQty:N0} pcs) has completed Hub {transition.CurrentHubName} ({transition.CurrentProcessName}) and is ready for Hub {nextHub} (Next process: {transition.NextFirstProcessName ?? "Next Step"}).";
                string actionUrl = $"/production/receive-lot?lotNo={lotNo}";

                string assignedToUserNames = string.Empty;
                string assignedToDisplayNames = string.Empty;

                if (transition.NextHubSupervisors.Any())
                {
                    assignedToUserNames = string.Join(",", transition.NextHubSupervisors.Select(s => s.UserName));
                    assignedToDisplayNames = string.Join(", ", transition.NextHubSupervisors.Select(s => s.EmployeeName ?? s.UserName));
                }
                else
                {
                    assignedToUserNames = "HubSupervisor";
                    assignedToDisplayNames = $"Hub {nextHub} Supervisors";
                }

                // 1. Insert Task in TaskItems
                const string insertTaskSql = @"
                    INSERT INTO TaskItems (
                        Title, Description, AssignedTo, AssignedToNames, AssignedBy, Priority, Status, DueDate,
                        WhatsAppMessageSent, EmailMessageSent, IsRead,
                        SourceEntityType, SourceEntityRefId, TargetRole, ActionUrl, CreatedAt, UpdatedAt
                    ) VALUES (
                        @Title, @Description, @AssignedTo, @AssignedToNames, @AssignedBy, 2, 0, DATEADD(day, 2, GETDATE()),
                        0, 0, 0,
                        'LotHub', @EntityRefId, 'HubSupervisor', @ActionUrl, GETUTCDATE(), GETUTCDATE()
                    );
                    SELECT CAST(SCOPE_IDENTITY() AS INT);";

                int taskId = await conn.ExecuteScalarAsync<int>(insertTaskSql, new
                {
                    Title = title,
                    Description = desc,
                    AssignedTo = assignedToUserNames,
                    AssignedToNames = assignedToDisplayNames,
                    AssignedBy = string.IsNullOrWhiteSpace(createdBy) ? "System" : createdBy,
                    EntityRefId = entityRefId,
                    ActionUrl = actionUrl
                }, trans);

                // 2. Dispatch notifications to all assigned supervisors and register in Task_Assignees
                int notifsCreated = 0;
                if (transition.NextHubSupervisors.Any())
                {
                    const string insertNotifSql = @"
                        INSERT INTO AppNotifications (
                            Id, Category, Title, Message, SenderName, TargetUserId, ActionUrl, CreatedAt, IsRead, IsReadReceipt
                        ) VALUES (
                            @Id, 4, @Title, @Message, @SenderName, @TargetUserId, @ActionUrl, GETUTCDATE(), 0, 0
                        );";

                    const string insertAssigneeSql = @"
                        IF NOT EXISTS (SELECT 1 FROM Task_Assignees WHERE TaskID = @TaskId AND UserID = @UserId)
                        BEGIN
                            INSERT INTO Task_Assignees (TaskID, UserID, UserName, AssignedAt)
                            VALUES (@TaskId, @UserId, @UserName, GETDATE())
                        END";

                    foreach (var supervisor in transition.NextHubSupervisors)
                    {
                        await conn.ExecuteAsync(insertNotifSql, new
                        {
                            Id = Guid.NewGuid().ToString(),
                            Title = title,
                            Message = desc,
                            SenderName = createdBy,
                            TargetUserId = supervisor.UserName,
                            ActionUrl = actionUrl
                        }, trans);

                        await conn.ExecuteAsync(insertAssigneeSql, new
                        {
                            TaskId = taskId,
                            UserId = supervisor.UserID,
                            UserName = supervisor.UserName
                        }, trans);

                        notifsCreated++;
                    }
                }

                trans.Commit();
                _logger.LogInformation("Created Task #{TaskId} and {NotifCount} notification(s) for Lot [{LotNo}] at Hub [{NextHub}]",
                    taskId, notifsCreated, lotNo, nextHub);

                return taskId;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating next hub task and notifications for Lot [{LotNo}]", lotNo);
                return 0;
            }
        }

        public async Task<int> SendDispatchFinalLotNotificationAsync(
            string lotNo,
            string itemCode,
            string orderNo,
            decimal rcvdQty,
            string completedProcessName,
            string createdBy)
        {
            try
            {
                using var conn = CreateConnection();
                await conn.OpenAsync();
                using var trans = conn.BeginTransaction();

                // 1. Fetch all active users with role 'Dispatch'
                const string usersSql = @"
                    SELECT DISTINCT u.UserID, u.UserName
                    FROM Users_User_Roles ur
                    INNER JOIN Users u ON ur.UserID = u.UserID
                    WHERE (ur.User_Role = 'Dispatch' OR ur.User_Role LIKE '%Dispatch%') 
                      AND COALESCE(u.InActive, 0) = 0";

                var dispatchUsers = (await conn.QueryAsync<dynamic>(usersSql, transaction: trans)).ToList();

                if (!dispatchUsers.Any())
                {
                    // Fallback to Administrator
                    const string fallbackSql = @"
                        SELECT UserID, UserName 
                        FROM Users 
                        WHERE COALESCE(InActive, 0) = 0 AND UserName = 'Administrator'";
                    dispatchUsers = (await conn.QueryAsync<dynamic>(fallbackSql, transaction: trans)).ToList();
                }

                if (!dispatchUsers.Any())
                {
                    _logger.LogWarning("No active users found for role 'Dispatch' to notify for Lot [{LotNo}] completion.", lotNo);
                    return 0;
                }

                string title = $"Lot #{lotNo} Ready for Dispatch (Order #{orderNo})";
                string message = $"Item [{itemCode}], Lot #{lotNo} ({rcvdQty:N0} pcs) has completed final process ({completedProcessName}) and is now available in 'Lots Finalized' for dispatch.";
                string entityRefId = $"{lotNo}:Dispatch";
                string assignedToUserNames = string.Join(",", dispatchUsers.Select(u => (string)u.UserName));
                string assignedToDisplayNames = "[Role: Dispatch]";

                // 2. Create TaskItem for Dispatch (ActionUrl = "" per request #10)
                const string insertTaskSql = @"
                    INSERT INTO TaskItems (
                        Title, Description, AssignedTo, AssignedToNames, AssignedBy, Priority, Status, DueDate,
                        WhatsAppMessageSent, EmailMessageSent, IsRead,
                        SourceEntityType, SourceEntityRefId, TargetRole, ActionUrl, CreatedAt, UpdatedAt
                    ) VALUES (
                        @Title, @Description, @AssignedTo, @AssignedToNames, @AssignedBy, 2, 0, DATEADD(day, 2, GETDATE()),
                        0, 0, 0,
                        'LotDispatch', @EntityRefId, 'Dispatch', '', GETUTCDATE(), GETUTCDATE()
                    );
                    SELECT CAST(SCOPE_IDENTITY() AS INT);";

                int taskId = await conn.ExecuteScalarAsync<int>(insertTaskSql, new
                {
                    Title = title,
                    Description = message,
                    AssignedTo = assignedToUserNames,
                    AssignedToNames = assignedToDisplayNames,
                    AssignedBy = string.IsNullOrWhiteSpace(createdBy) ? "System" : createdBy,
                    EntityRefId = entityRefId
                }, trans);

                // 3. Register in Task_Assignees & AppNotifications
                const string insertNotifSql = @"
                    INSERT INTO AppNotifications (
                        Id, Category, Title, Message, SenderName, TargetUserId, ActionUrl, CreatedAt, IsRead, IsReadReceipt
                    ) VALUES (
                        @Id, 4, @Title, @Message, @SenderName, @TargetUserId, '', GETUTCDATE(), 0, 0
                    );";

                const string insertAssigneeSql = @"
                    IF NOT EXISTS (SELECT 1 FROM Task_Assignees WHERE TaskID = @TaskId AND UserID = @UserId)
                    BEGIN
                        INSERT INTO Task_Assignees (TaskID, UserID, UserName, AssignedAt)
                        VALUES (@TaskId, @UserId, @UserName, GETDATE())
                    END";

                int count = 0;
                foreach (var user in dispatchUsers)
                {
                    int uId = (int)user.UserID;
                    string uName = (string)user.UserName;

                    await conn.ExecuteAsync(insertAssigneeSql, new
                    {
                        TaskId = taskId,
                        UserId = uId,
                        UserName = uName
                    }, trans);

                    await conn.ExecuteAsync(insertNotifSql, new
                    {
                        Id = Guid.NewGuid().ToString(),
                        Title = title,
                        Message = message,
                        SenderName = createdBy,
                        TargetUserId = uName
                    }, trans);

                    count++;
                }

                trans.Commit();
                _logger.LogInformation("Created Task #{TaskId} and dispatched {Count} notification(s) to 'Dispatch' role for Lot [{LotNo}] completion.", taskId, count, lotNo);
                return count;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error sending dispatch notification for Lot [{LotNo}]", lotNo);
                return 0;
            }
        }
    }
}
