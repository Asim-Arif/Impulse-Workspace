using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.IntraOffice;
using DataAccessLibrary.Interface.Setup;
using DataAccessLibrary.Models.IntraOffice;
using Impulse.Services.IntraOffice;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.WorkflowTasks
{
    public class WorkflowTaskEngine : IWorkflowTaskEngine
    {
        private readonly IIntraOfficeDataAccess _intraData;
        private readonly IUserRoleDataAccess _userRoleData;
        private readonly IAppNotificationService _notificationService;
        private readonly string _connectionString;
        private readonly ILogger<WorkflowTaskEngine> _logger;

        public WorkflowTaskEngine(
            IIntraOfficeDataAccess intraData,
            IUserRoleDataAccess userRoleData,
            IAppNotificationService notificationService,
            IConfiguration configuration,
            ILogger<WorkflowTaskEngine> logger)
        {
            _intraData = intraData;
            _userRoleData = userRoleData;
            _notificationService = notificationService;
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task<int> CreateRoleTaskAsync(WorkflowTaskCreateRequest request)
        {
            try
            {
                var roles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (!string.IsNullOrWhiteSpace(request.TargetRole))
                {
                    roles.Add(request.TargetRole.Trim());
                }
                if (request.TargetRoles != null)
                {
                    foreach (var r in request.TargetRoles)
                    {
                        if (!string.IsNullOrWhiteSpace(r)) roles.Add(r.Trim());
                    }
                }

                var rolesList = roles.ToList();
                var primaryRole = rolesList.FirstOrDefault() ?? request.TargetRole;
                var roleCaption = rolesList.Count > 0 ? string.Join(" & ", rolesList) : "Assigned Role";

                // Resolve target users for notifications and assignees
                var targetUserNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var role in rolesList)
                {
                    var usersInRole = await _userRoleData.GetUsersByRoleAsync(role);
                    foreach (var u in usersInRole)
                    {
                        if (!u.InActive && !string.IsNullOrWhiteSpace(u.UserName))
                        {
                            targetUserNames.Add(u.UserName);
                        }
                    }
                }

                if (request.TargetUserNames != null)
                {
                    foreach (var u in request.TargetUserNames)
                    {
                        if (!string.IsNullOrWhiteSpace(u)) targetUserNames.Add(u.Trim());
                    }
                }

                var task = new TaskItem
                {
                    Title = request.Title,
                    Description = request.Description,
                    SourceEntityType = request.SourceEntityType,
                    SourceEntityRefId = request.SourceEntityRefId,
                    TargetRole = primaryRole,
                    TargetRoles = rolesList,
                    ActionUrl = request.ActionUrl,
                    Priority = (TaskPriority)request.Priority,
                    Status = TaskItemStatus.Pending,
                    DueDate = request.DueDate,
                    AssignedBy = request.CreatedBy,
                    AssignedTo = primaryRole,
                    AdditionalAssigneeIds = targetUserNames.Count > 0 ? string.Join(",", targetUserNames) : null,
                    AssignedToNames = targetUserNames.Count > 0 && string.Equals(primaryRole, "HubSupervisor", StringComparison.OrdinalIgnoreCase)
                        ? $"[Hub Supervisors: {string.Join(", ", targetUserNames)}]"
                        : $"[Role: {roleCaption}]"
                };

                // Save task and child table roles
                var taskId = await _intraData.CreateTaskAsync(task);
                task.Id = taskId;

                if (targetUserNames.Count == 0)
                {
                    _logger.LogWarning("No target users resolved for role '{Role}' or explicit assignees on Task #{TaskId}. Notification table entry skipped.",
                        primaryRole, taskId);
                }

                // Dispatch real-time pop-up notification & bell alert to all target users and persist directly to AppNotifications
                foreach (var userName in targetUserNames)
                {
                    try
                    {
                        await _notificationService.SendNotificationAsync(new AppNotification
                        {
                            Category = NotificationCategory.Task,
                            Title = request.Title,
                            Message = request.Description,
                            SenderName = request.CreatedBy,
                            TargetUserId = userName,
                            ActionUrl = request.ActionUrl
                        });

                        _logger.LogInformation("AppNotification successfully generated for user '{User}' on Task #{TaskId}", userName, taskId);
                    }
                    catch (Exception notifEx)
                    {
                        _logger.LogError(notifEx, "Failed to persist AppNotification for user '{User}' on Task #{TaskId}", userName, taskId);
                    }
                }

                _logger.LogInformation("Workflow task #{TaskId} created for entity {EntityType} #{EntityRef} targeting {RoleCount} roles and {UserCount} users",
                    taskId, request.SourceEntityType, request.SourceEntityRefId, rolesList.Count, targetUserNames.Count);

                return taskId;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating role workflow task for {EntityType} #{EntityRef}",
                    request.SourceEntityType, request.SourceEntityRefId);
                return 0;
            }
        }

        public async Task<bool> CompleteTaskAsync(string entityType, string entityRefId, string completedByUserName, string? notes = null)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    UPDATE TaskItems 
                    SET Status = 2, 
                        CompletedAt = GETUTCDATE(), 
                        CompletedBy = @CompletedBy, 
                        UpdatedAt = GETUTCDATE()
                    WHERE SourceEntityType = @EntityType 
                      AND SourceEntityRefId = @EntityRefId 
                      AND Status <> 2";

                var rows = await db.ExecuteAsync(sql, new
                {
                    EntityType = entityType,
                    EntityRefId = entityRefId,
                    CompletedBy = completedByUserName
                });

                if (!string.IsNullOrWhiteSpace(notes) && rows > 0)
                {
                    var commentSql = @"
                        INSERT INTO TaskComments (TaskId, UserId, Content, CreatedAt)
                        SELECT Id, @CompletedBy, @Content, GETUTCDATE()
                        FROM TaskItems
                        WHERE SourceEntityType = @EntityType AND SourceEntityRefId = @EntityRefId";
                    await db.ExecuteAsync(commentSql, new
                    {
                        EntityType = entityType,
                        EntityRefId = entityRefId,
                        CompletedBy = completedByUserName,
                        Content = notes
                    });
                }

                return rows > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error completing task for {EntityType} #{EntityRef}", entityType, entityRefId);
                return false;
            }
        }

        public async Task<bool> CompleteRoleTaskAsync(string entityType, string entityRefId, string targetRole, string completedByUserName, string? notes = null)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    UPDATE TaskItems 
                    SET Status = 2, 
                        CompletedAt = GETUTCDATE(), 
                        CompletedBy = @CompletedBy, 
                        UpdatedAt = GETUTCDATE()
                    WHERE SourceEntityType = @EntityType 
                      AND SourceEntityRefId = @EntityRefId 
                      AND (TargetRole = @TargetRole OR AssignedTo = @TargetRole)
                      AND Status <> 2";

                var rows = await db.ExecuteAsync(sql, new
                {
                    EntityType = entityType,
                    EntityRefId = entityRefId,
                    TargetRole = targetRole,
                    CompletedBy = completedByUserName
                });

                if (!string.IsNullOrWhiteSpace(notes) && rows > 0)
                {
                    var commentSql = @"
                        INSERT INTO TaskComments (TaskId, UserId, Content, CreatedAt)
                        SELECT Id, @CompletedBy, @Content, GETUTCDATE()
                        FROM TaskItems
                        WHERE SourceEntityType = @EntityType 
                          AND SourceEntityRefId = @EntityRefId
                          AND (TargetRole = @TargetRole OR AssignedTo = @TargetRole)";
                    await db.ExecuteAsync(commentSql, new
                    {
                        EntityType = entityType,
                        EntityRefId = entityRefId,
                        TargetRole = targetRole,
                        CompletedBy = completedByUserName,
                        Content = notes
                    });
                }

                if (rows > 0)
                {
                    var notifSql = @"
                        UPDATE AppNotifications 
                        SET IsRead = 1, ReadAt = GETUTCDATE()
                        WHERE TaskId IN (
                            SELECT Id FROM TaskItems 
                            WHERE SourceEntityType = @EntityType 
                              AND SourceEntityRefId = @EntityRefId
                              AND (TargetRole = @TargetRole OR AssignedTo = @TargetRole)
                        ) AND IsRead = 0";
                    await db.ExecuteAsync(notifSql, new
                    {
                        EntityType = entityType,
                        EntityRefId = entityRefId,
                        TargetRole = targetRole
                    });
                }

                return rows > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error completing {Role} task for {EntityType} #{EntityRef}", targetRole, entityType, entityRefId);
                return false;
            }
        }

        public async Task<bool> AuthorizeCustomerOrderAsync(string orderNo, string authorizedByUserName)
        {
            try
            {
                using var db = CreateConnection();
                db.Open();
                using var trans = db.BeginTransaction();

                // 1. Authorize Customer Order
                var updateOrderSql = @"
                    UPDATE FCustomerOrders 
                    SET Authorized = 1, 
                        AuthorizedBy = @AuthorizedBy, 
                        AuthorizedDT = GETDATE()
                    WHERE OrderNo = @OrderNo";

                var affected = await db.ExecuteAsync(updateOrderSql, new
                {
                    OrderNo = orderNo,
                    AuthorizedBy = authorizedByUserName
                }, trans);

                if (affected == 0)
                {
                    trans.Rollback();
                    return false;
                }

                // 2. Mark Director task completed
                var completeTaskSql = @"
                    UPDATE TaskItems 
                    SET Status = 2, 
                        CompletedAt = GETUTCDATE(), 
                        CompletedBy = @AuthorizedBy, 
                        UpdatedAt = GETUTCDATE()
                    WHERE SourceEntityType = 'CustomerOrder' 
                      AND SourceEntityRefId = @OrderNo 
                      AND Status <> 2";

                await db.ExecuteAsync(completeTaskSql, new
                {
                    OrderNo = orderNo,
                    AuthorizedBy = authorizedByUserName
                }, trans);

                // 3. Log comment
                var commentSql = @"
                    INSERT INTO TaskComments (TaskId, UserId, Content, CreatedAt)
                    SELECT Id, @AuthorizedBy, 'Order Authorized by Director. Automatic handover to PPC initiated.', GETUTCDATE()
                    FROM TaskItems
                    WHERE SourceEntityType = 'CustomerOrder' AND SourceEntityRefId = @OrderNo";

                await db.ExecuteAsync(commentSql, new
                {
                    OrderNo = orderNo,
                    AuthorizedBy = authorizedByUserName
                }, trans);

                trans.Commit();

                _logger.LogInformation("Customer Order #{OrderNo} successfully authorized by {User}. Triggering PPC workflow task.", orderNo, authorizedByUserName);

                // 4. Automatically generate handover task for PPC role
                _ = Task.Run(async () =>
                {
                    try
                    {
                        await CreateRoleTaskAsync(new WorkflowTaskCreateRequest
                        {
                            SourceEntityType = "CustomerOrder",
                            SourceEntityRefId = orderNo,
                            TargetRole = "PPC",
                            Title = $"PPC Planning for Authorized Order #{orderNo}",
                            Description = $"Customer Order #{orderNo} has been authorized by Director ({authorizedByUserName}). Ready for material planning, store issuance, and production orders.",
                            ActionUrl = $"/production/order-planning?orderNo={orderNo}",
                            Priority = 1,
                            DueDate = DateTime.Today.AddDays(2),
                            CreatedBy = authorizedByUserName
                        });
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Error creating follow-up PPC task for Customer Order #{OrderNo}", orderNo);
                    }
                });

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error authorizing customer order #{OrderNo}", orderNo);
                return false;
            }
        }
    }
}
