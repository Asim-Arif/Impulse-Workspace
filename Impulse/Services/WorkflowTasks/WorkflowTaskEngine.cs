using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.IntraOffice;
using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Interface.Setup;
using DataAccessLibrary.Models.IntraOffice;
using DataAccessLibrary.Models.ViewModels.Payroll;
using DataAccessLibrary.Models.ViewModels.Production;
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
        private readonly IUserDataAccess _userData;
        private readonly IWorkflowConfigurationDataAccess _workflowConfigData;
        private readonly IAppNotificationService _notificationService;
        private readonly IAccountReportingAccess _accountReportingAccess;
        private readonly IManualAttendanceDataAccess _manualAttendanceData;
        private readonly IMonthlyAttendanceDataAccess _monthlyAttendanceData;
        private readonly IMakerItemAssignmentDataAccess _makerItemData;
        private readonly IProductionDeletionDataAccess _prodDeletionData;
        private readonly IMakerRcvListDataAccess _makerRcvListData;
        private readonly IMakerPOListDataAccess _makerPOListData;
        private readonly ILotIssuanceDataAccess _lotIssuanceData;
        private readonly string _connectionString;
        private readonly ILogger<WorkflowTaskEngine> _logger;

        public WorkflowTaskEngine(
            IIntraOfficeDataAccess intraData,
            IUserRoleDataAccess userRoleData,
            IUserDataAccess userData,
            IWorkflowConfigurationDataAccess workflowConfigData,
            IAppNotificationService notificationService,
            IAccountReportingAccess accountReportingAccess,
            IManualAttendanceDataAccess manualAttendanceData,
            IMonthlyAttendanceDataAccess monthlyAttendanceData,
            IMakerItemAssignmentDataAccess makerItemData,
            IProductionDeletionDataAccess prodDeletionData,
            IMakerRcvListDataAccess makerRcvListData,
            IMakerPOListDataAccess makerPOListData,
            ILotIssuanceDataAccess lotIssuanceData,
            IConfiguration configuration,
            ILogger<WorkflowTaskEngine> logger)
        {
            _intraData = intraData;
            _userRoleData = userRoleData;
            _userData = userData;
            _workflowConfigData = workflowConfigData;
            _notificationService = notificationService;
            _accountReportingAccess = accountReportingAccess;
            _manualAttendanceData = manualAttendanceData;
            _monthlyAttendanceData = monthlyAttendanceData;
            _makerItemData = makerItemData;
            _prodDeletionData = prodDeletionData;
            _makerRcvListData = makerRcvListData;
            _makerPOListData = makerPOListData;
            _lotIssuanceData = lotIssuanceData;
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

                // Idempotency check: prevent duplicate active (Pending=0, InProgress=1) tasks for the same entity and role
                if (!string.IsNullOrWhiteSpace(request.SourceEntityType) && !string.IsNullOrWhiteSpace(request.SourceEntityRefId))
                {
                    using var checkDb = CreateConnection();
                    const string existingSql = @"
                        SELECT TOP 1 Id 
                        FROM TaskItems 
                        WHERE SourceEntityType = @SourceEntityType 
                          AND SourceEntityRefId = @SourceEntityRefId 
                          AND TargetRole = @TargetRole 
                          AND Status IN (0, 1)";

                    var existingId = await checkDb.ExecuteScalarAsync<int?>(existingSql, new
                    {
                        request.SourceEntityType,
                        request.SourceEntityRefId,
                        TargetRole = primaryRole
                    });

                    if (existingId.HasValue && existingId.Value > 0)
                    {
                        _logger.LogInformation("Active workflow task #{TaskId} already exists for {EntityType} #{EntityRef} ({Role}). Skipping duplicate creation and notifications.",
                            existingId.Value, request.SourceEntityType, request.SourceEntityRefId, primaryRole);
                        return existingId.Value;
                    }
                }

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

                // 1. Check pending items / deltas before updating
                var pendingItemsSql = @"
                    SELECT oi.ID, oi.CompItemCode, oi.ItemCode, oi.Qty, ISNULL(oi.AuthorizedQty, 0) AS AuthorizedQty,
                           ISNULL(v.ItemName, oi.CompItemCode) AS ItemName,
                           (oi.Qty - ISNULL(oi.AuthorizedQty, 0)) AS DeltaQty
                    FROM FOrderItems oi
                    LEFT JOIN VItems v ON oi.CompItemCode = v.ItemID
                    WHERE oi.OrderNo = @OrderNo AND (ISNULL(oi.Authorized, 0) = 0 OR oi.Qty <> ISNULL(oi.AuthorizedQty, 0))";

                var pendingItems = (await db.QueryAsync<dynamic>(pendingItemsSql, new { OrderNo = orderNo }, trans)).ToList();

                var previouslyAuthCountSql = @"
                    SELECT COUNT(*) FROM FOrderItems 
                    WHERE OrderNo = @OrderNo AND ISNULL(Authorized, 0) = 1 AND ISNULL(AuthorizedQty, 0) > 0";
                int previouslyAuthCount = await db.QueryFirstOrDefaultAsync<int>(previouslyAuthCountSql, new { OrderNo = orderNo }, trans);
                bool isReauthorization = previouslyAuthCount > 0 && pendingItems.Any();

                // 2. Authorize pending item lines in FOrderItems
                var updateItemsSql = @"
                    UPDATE FOrderItems 
                    SET Authorized = 1,
                        AuthorizedQty = Qty,
                        AuthorizedBy = @AuthorizedBy,
                        AuthorizedDT = GETDATE()
                    WHERE OrderNo = @OrderNo AND (ISNULL(Authorized, 0) = 0 OR Qty <> ISNULL(AuthorizedQty, 0))";

                await db.ExecuteAsync(updateItemsSql, new
                {
                    OrderNo = orderNo,
                    AuthorizedBy = authorizedByUserName
                }, trans);

                // 3. Authorize Customer Order Header
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

                // 4. Mark Director tasks completed
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

                // 5. Log comment
                var commentText = isReauthorization
                    ? $"Order Re-Authorized by Director ({authorizedByUserName}). {pendingItems.Count} new/modified item(s) approved and handed over to PPC."
                    : $"Order Authorized by Director ({authorizedByUserName}). Automatic handover to PPC initiated.";

                var commentSql = @"
                    INSERT INTO TaskComments (TaskId, UserId, Content, CreatedAt)
                    SELECT Id, @AuthorizedBy, @CommentText, GETUTCDATE()
                    FROM TaskItems
                    WHERE SourceEntityType = 'CustomerOrder' AND SourceEntityRefId = @OrderNo";

                await db.ExecuteAsync(commentSql, new
                {
                    OrderNo = orderNo,
                    AuthorizedBy = authorizedByUserName,
                    CommentText = commentText
                }, trans);

                trans.Commit();

                _logger.LogInformation("Customer Order #{OrderNo} successfully authorized by {User}. Triggering PPC workflow task.", orderNo, authorizedByUserName);

                // 6. Automatically generate handover task for PPC role
                var itemSummaryList = pendingItems.Select(p => 
                    (int)p.AuthorizedQty > 0 
                        ? $"{p.ItemName} (Qty: {p.Qty}, Delta: +{p.DeltaQty})" 
                        : $"{p.ItemName} (Qty: {p.Qty})").ToList();

                string itemsDetails = itemSummaryList.Any() 
                    ? string.Join(", ", itemSummaryList.Take(5)) + (itemSummaryList.Count > 5 ? $" (+{itemSummaryList.Count - 5} more)" : "")
                    : "all items";

                _ = Task.Run(async () =>
                {
                    try
                    {
                        string taskTitle = isReauthorization
                            ? $"[Update] New Items Authorized for Order #{orderNo}"
                            : $"PPC Planning for Authorized Order #{orderNo}";

                        string taskDesc = isReauthorization
                            ? $"Director ({authorizedByUserName}) authorized new/modified items for Order #{orderNo}: {itemsDetails}. Please complete material planning, store issuance, and production orders for these items."
                            : $"Customer Order #{orderNo} has been authorized by Director ({authorizedByUserName}). Ready for material planning, store issuance, and production orders.";

                        await CreateRoleTaskAsync(new WorkflowTaskCreateRequest
                        {
                            SourceEntityType = "CustomerOrder",
                            SourceEntityRefId = orderNo,
                            TargetRole = "PPC",
                            Title = taskTitle,
                            Description = taskDesc,
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
                throw;
            }
        }

        public async Task<bool> RejectRoleTaskAsync(string entityType, string entityRefId, string targetRole, string rejectedByUserName, string? rejectionReason = null)
        {
            try
            {
                using var db = CreateConnection();
                var sql = @"
                    UPDATE TaskItems 
                    SET Status = 3, 
                        CompletedBy = @RejectedBy, 
                        UpdatedAt = GETUTCDATE()
                    WHERE SourceEntityType = @EntityType 
                      AND SourceEntityRefId = @EntityRefId 
                      AND (TargetRole = @TargetRole OR AssignedTo = @TargetRole)
                      AND Status IN (0, 1)";

                var rows = await db.ExecuteAsync(sql, new
                {
                    EntityType = entityType,
                    EntityRefId = entityRefId,
                    TargetRole = targetRole,
                    RejectedBy = rejectedByUserName
                });

                if (!string.IsNullOrWhiteSpace(rejectionReason) && rows > 0)
                {
                    var commentSql = @"
                        INSERT INTO TaskComments (TaskId, UserId, Content, CreatedAt)
                        SELECT Id, @RejectedBy, @Content, GETUTCDATE()
                        FROM TaskItems
                        WHERE SourceEntityType = @EntityType 
                          AND SourceEntityRefId = @EntityRefId
                          AND (TargetRole = @TargetRole OR AssignedTo = @TargetRole)";
                    await db.ExecuteAsync(commentSql, new
                    {
                        EntityType = entityType,
                        EntityRefId = entityRefId,
                        TargetRole = targetRole,
                        RejectedBy = rejectedByUserName,
                        Content = rejectionReason
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
                _logger.LogError(ex, "Error rejecting {Role} task for {EntityType} #{EntityRef}", targetRole, entityType, entityRefId);
                return false;
            }
        }

        public async Task<int> RequestVoucherDeletionAsync(string vchrNo, string originatorUserName, string deleteReason, string? machineName = null)
        {
            try
            {
                using var db = CreateConnection();
                
                // Idempotency check: prevent duplicate active deletion tasks for the same voucher
                const string checkSql = @"
                    SELECT TOP 1 Id FROM TaskItems 
                    WHERE SourceEntityType = 'VoucherDeletion' 
                      AND SourceEntityRefId = @VchrNo 
                      AND Status IN (0, 1)";
                var existingTaskId = await db.ExecuteScalarAsync<int?>(checkSql, new { VchrNo = vchrNo });
                if (existingTaskId.HasValue && existingTaskId.Value > 0)
                {
                    _logger.LogInformation("Active deletion task #{TaskId} already exists for voucher #{VchrNo}", existingTaskId.Value, vchrNo);
                    return existingTaskId.Value;
                }

                // Fetch voucher summary (Date & total amount)
                const string voucherSummarySql = @"
                    SELECT TOP 1 VDate, Description, 
                           (SELECT ISNULL(SUM(Debit), 0) FROM Vouchers WHERE VchrNo = @VchrNo) AS TotalAmount
                    FROM Vouchers 
                    WHERE VchrNo = @VchrNo";
                var voucherSummary = await db.QueryFirstOrDefaultAsync<dynamic>(voucherSummarySql, new { VchrNo = vchrNo });

                DateTime vDate = voucherSummary != null ? Convert.ToDateTime(voucherSummary.VDate) : DateTime.Today;
                decimal totalAmount = voucherSummary != null ? Convert.ToDecimal(voucherSummary.TotalAmount) : 0m;

                string title = $"⚠️ Voucher Deletion Request: #{vchrNo}";
                string desc = $"Requested by {originatorUserName} | Reason: {deleteReason} | Amount: PKR {totalAmount:N2} | Date: {vDate:yyyy-MM-dd}";

                var taskId = await CreateRoleTaskAsync(new WorkflowTaskCreateRequest
                {
                    SourceEntityType = "VoucherDeletion",
                    SourceEntityRefId = vchrNo,
                    TargetRole = "Director",
                    Title = title,
                    Description = desc,
                    ActionUrl = $"/accounts/voucher-approval?vchrNo={vchrNo}",
                    Priority = 3, // Urgent
                    DueDate = DateTime.Today.AddDays(1),
                    CreatedBy = originatorUserName
                });

                if (taskId > 0 && !string.IsNullOrWhiteSpace(deleteReason))
                {
                    var commentSql = @"
                        INSERT INTO TaskComments (TaskId, UserId, Content, CreatedAt)
                        VALUES (@TaskId, @UserId, @Content, GETUTCDATE())";
                    await db.ExecuteAsync(commentSql, new
                    {
                        TaskId = taskId,
                        UserId = originatorUserName,
                        Content = $"[Deletion Reason] {deleteReason}" + (!string.IsNullOrWhiteSpace(machineName) ? $" (Station: {machineName})" : "")
                    });
                }

                return taskId;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error requesting voucher deletion task for Voucher #{VchrNo}", vchrNo);
                throw;
            }
        }

        public async Task<bool> ApproveVoucherDeletionAsync(string vchrNo, string approvedByDirectorUserName, string? directorRemarks = null)
        {
            try
            {
                using var db = CreateConnection();
                // 1. Fetch active task for this voucher
                const string taskSql = @"
                    SELECT TOP 1 * FROM TaskItems 
                    WHERE SourceEntityType = 'VoucherDeletion' 
                      AND SourceEntityRefId = @VchrNo 
                      AND Status IN (0, 1)
                    ORDER BY Id DESC";
                var task = await db.QueryFirstOrDefaultAsync<TaskItem>(taskSql, new { VchrNo = vchrNo });

                string originator = task?.AssignedBy ?? "Unknown";

                // Fetch originator deletion reason from comments
                const string commentSql = @"
                    SELECT TOP 1 Content FROM TaskComments 
                    WHERE TaskId = @TaskId 
                    ORDER BY Id ASC";
                string? storedComment = task != null ? await db.ExecuteScalarAsync<string>(commentSql, new { TaskId = task.Id }) : null;
                string reason = !string.IsNullOrWhiteSpace(storedComment) ? storedComment : task?.Description ?? "Deletion Approved by Director";

                // 2. Call existing AccountReportingAccess.DeleteVoucher for 100% full business deletion
                var model = new DataAccessLibrary.Models.ViewModels.Accounts.AccountsReportingModel
                {
                    VchrNo = vchrNo,
                    DeleteReason = $"[Approved by Director {approvedByDirectorUserName}] {reason}" + (!string.IsNullOrWhiteSpace(directorRemarks) ? $" | Remarks: {directorRemarks}" : ""),
                    UserName = originator,
                    MachineName = Environment.MachineName
                };

                await _accountReportingAccess.DeleteVoucher(model, vchrNo, false);

                // 3. Mark role task completed
                if (task != null)
                {
                    await CompleteRoleTaskAsync("VoucherDeletion", vchrNo, "Director", approvedByDirectorUserName, 
                        $"Voucher deletion approved and permanently removed by Director {approvedByDirectorUserName}." + 
                        (!string.IsNullOrWhiteSpace(directorRemarks) ? $" Remarks: {directorRemarks}" : ""));
                }

                // 4. Send real-time notification + web push back to originator
                if (!string.IsNullOrWhiteSpace(originator) && !string.Equals(originator, "System", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        await _notificationService.SendNotificationAsync(new AppNotification
                        {
                            Category = NotificationCategory.Task,
                            Title = $"✅ Voucher #{vchrNo} Deletion Approved",
                            Message = $"Your deletion request for Voucher #{vchrNo} was approved and deleted by Director {approvedByDirectorUserName}." + 
                                      (!string.IsNullOrWhiteSpace(directorRemarks) ? $" Remarks: {directorRemarks}" : ""),
                            SenderName = approvedByDirectorUserName,
                            TargetUserId = originator,
                            ActionUrl = "/accounts/transactionregister"
                        });
                    }
                    catch (Exception notifEx)
                    {
                        _logger.LogError(notifEx, "Failed to send approval notification to originator '{Originator}' for Voucher #{VchrNo}", originator, vchrNo);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error approving voucher deletion for Voucher #{VchrNo}", vchrNo);
                throw;
            }
        }

        public async Task<bool> RejectVoucherDeletionAsync(string vchrNo, string rejectedByDirectorUserName, string rejectionReason)
        {
            try
            {
                using var db = CreateConnection();
                // 1. Fetch active task
                const string taskSql = @"
                    SELECT TOP 1 * FROM TaskItems 
                    WHERE SourceEntityType = 'VoucherDeletion' 
                      AND SourceEntityRefId = @VchrNo 
                      AND Status IN (0, 1)
                    ORDER BY Id DESC";
                var task = await db.QueryFirstOrDefaultAsync<TaskItem>(taskSql, new { VchrNo = vchrNo });

                string originator = task?.AssignedBy ?? "Unknown";

                // 2. Reject the role task
                await RejectRoleTaskAsync("VoucherDeletion", vchrNo, "Director", rejectedByDirectorUserName, 
                    $"Voucher deletion request rejected by Director {rejectedByDirectorUserName}. Reason: {rejectionReason}");

                // 3. Send real-time notification + web push back to originator
                if (!string.IsNullOrWhiteSpace(originator) && !string.Equals(originator, "System", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        await _notificationService.SendNotificationAsync(new AppNotification
                        {
                            Category = NotificationCategory.Task,
                            Title = $"❌ Voucher #{vchrNo} Deletion Rejected",
                            Message = $"Your deletion request for Voucher #{vchrNo} was rejected by Director {rejectedByDirectorUserName}. Remarks: {rejectionReason}",
                            SenderName = rejectedByDirectorUserName,
                            TargetUserId = originator,
                            ActionUrl = $"/accounts/transactionregister?p_VchrNo={vchrNo}"
                        });
                    }
                    catch (Exception notifEx)
                    {
                        _logger.LogError(notifEx, "Failed to send rejection notification to originator '{Originator}' for Voucher #{VchrNo}", originator, vchrNo);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error rejecting voucher deletion for Voucher #{VchrNo}", vchrNo);
                throw;
            }
        }

        public async Task<List<TaskItem>> GetPendingVoucherDeletionTasksAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT t.*, 
                           c.Content AS Description
                    FROM TaskItems t
                    OUTER APPLY (
                        SELECT TOP 1 Content 
                        FROM TaskComments 
                        WHERE TaskId = t.Id 
                        ORDER BY Id ASC
                    ) c
                    WHERE t.SourceEntityType = 'VoucherDeletion' 
                      AND t.Status IN (0, 1)
                    ORDER BY t.CreatedAt DESC";
                var items = await db.QueryAsync<TaskItem>(sql);
                return items.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching pending voucher deletion tasks");
                return new List<TaskItem>();
            }
        }

        #region Attendance Workflow Methods

        public async Task<int> RequestAttendanceActionAsync(AttendanceWorkflowRequestDto request)
        {
            try
            {
                // Month-wide salary finalization safeguard check
                if (await _monthlyAttendanceData.IsSalaryFinalizedAsync(request.Year, request.Month))
                {
                    throw new InvalidOperationException($"Salary for {request.Year}-{request.Month:D2} has already been finalized. Attendance modification request cannot be created.");
                }

                string entityRefId = $"{request.ActionType}_{request.EmpID}_{request.Year:D4}{request.Month:D2}" +
                                     (request.AttendanceDate.HasValue ? $"_{request.AttendanceDate.Value:dd}" : "");

                using var db = CreateConnection();

                // Check for existing active request
                const string checkSql = @"
                    SELECT TOP 1 Id FROM TaskItems 
                    WHERE SourceEntityType = 'AttendanceWorkflow' 
                      AND SourceEntityRefId = @RefId 
                      AND Status IN (0, 1)";
                var existingTaskId = await db.ExecuteScalarAsync<int?>(checkSql, new { RefId = entityRefId });
                if (existingTaskId.HasValue && existingTaskId.Value > 0)
                {
                    _logger.LogInformation("Active attendance workflow task #{TaskId} already exists for ref {RefId}", existingTaskId.Value, entityRefId);
                    return existingTaskId.Value;
                }

                string actionLabel = request.ActionType switch
                {
                    AttendanceWorkflowActionType.ManualSave => "Manual Attendance Edit",
                    AttendanceWorkflowActionType.ManualDelete => "Manual Attendance Deletion",
                    AttendanceWorkflowActionType.MonthlySave => "Monthly Attendance Batch Edit",
                    AttendanceWorkflowActionType.MonthlyClearDate => "Clear Date Attendance",
                    _ => "Attendance Modification"
                };

                string dateLabel = request.AttendanceDate.HasValue
                    ? request.AttendanceDate.Value.ToString("dd-MMM-yyyy")
                    : $"{request.Year}-{request.Month:D2}";

                string title = $"🕒 Attendance Approval: [{request.EmpID}] {request.EmployeeName} - {actionLabel}";
                string desc = $"Employee: [{request.EmpID}] {request.EmployeeName} | Period: {dateLabel} | Reason: {request.Reason} | Requested by: {request.OriginatorUserName}";

                var taskId = await CreateRoleTaskAsync(new WorkflowTaskCreateRequest
                {
                    SourceEntityType = "AttendanceWorkflow",
                    SourceEntityRefId = entityRefId,
                    TargetRole = "Director",
                    Title = title,
                    Description = desc,
                    ActionUrl = "/payroll/attendance-approval",
                    Priority = 2,
                    DueDate = DateTime.Today.AddDays(2),
                    CreatedBy = request.OriginatorUserName
                });

                request.TaskId = taskId;
                string payloadJson = JsonSerializer.Serialize(request);

                // Store serialized payload in TaskComments for lossless replay
                const string commentSql = @"
                    INSERT INTO TaskComments (TaskId, UserId, Content, CreatedAt)
                    VALUES (@TaskId, @UserId, @Content, GETUTCDATE())";
                await db.ExecuteAsync(commentSql, new
                {
                    TaskId = taskId,
                    UserId = request.OriginatorUserName,
                    Content = payloadJson
                });

                return taskId;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error requesting attendance workflow task for EmpID {EmpID}", request.EmpID);
                throw;
            }
        }

        public async Task<bool> ApproveAttendanceActionAsync(int taskId, string approvedByDirectorUserName, string? directorRemarks = null)
        {
            try
            {
                var request = await GetAttendanceTaskDetailsAsync(taskId);
                if (request == null)
                {
                    throw new InvalidOperationException($"Attendance task #{taskId} details could not be found.");
                }

                // Month-wide salary finalization safeguard check
                if (await _monthlyAttendanceData.IsSalaryFinalizedAsync(request.Year, request.Month))
                {
                    throw new InvalidOperationException($"Approval Blocked: Salary for {request.Year}-{request.Month:D2} has already been finalized in MonthlySalaries. Attendance changes cannot be applied.");
                }

                // Execute the requested business operation
                switch (request.ActionType)
                {
                    case AttendanceWorkflowActionType.ManualSave:
                        if (request.ManualInput != null)
                        {
                            await _manualAttendanceData.SaveManualAttendanceAsync(request.ManualInput);
                        }
                        break;

                    case AttendanceWorkflowActionType.ManualDelete:
                        if (request.AttendanceDate.HasValue)
                        {
                            await _manualAttendanceData.DeleteAttendanceAsync(request.EmpID, request.AttendanceDate.Value);
                        }
                        break;

                    case AttendanceWorkflowActionType.MonthlySave:
                        if (request.MonthlyInput != null)
                        {
                            await _monthlyAttendanceData.SaveMonthlyAttendanceAsync(request.MonthlyInput);
                        }
                        break;

                    case AttendanceWorkflowActionType.MonthlyClearDate:
                        if (request.AttendanceDate.HasValue)
                        {
                            await _monthlyAttendanceData.ClearDateAttendanceAsync(request.EmpID, request.AttendanceDate.Value);
                        }
                        break;
                }

                string entityRefId = $"{request.ActionType}_{request.EmpID}_{request.Year:D4}{request.Month:D2}" +
                                     (request.AttendanceDate.HasValue ? $"_{request.AttendanceDate.Value:dd}" : "");

                // Mark task complete
                await CompleteRoleTaskAsync("AttendanceWorkflow", entityRefId, "Director", approvedByDirectorUserName,
                    $"Attendance change approved and applied by Director {approvedByDirectorUserName}." +
                    (!string.IsNullOrWhiteSpace(directorRemarks) ? $" Remarks: {directorRemarks}" : ""));

                // Notify originator
                if (!string.IsNullOrWhiteSpace(request.OriginatorUserName) && !string.Equals(request.OriginatorUserName, "System", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        await _notificationService.SendNotificationAsync(new AppNotification
                        {
                            Category = NotificationCategory.Task,
                            Title = $"✅ Attendance Approved: [{request.EmpID}] {request.EmployeeName}",
                            Message = $"Your attendance change request for [{request.EmpID}] {request.EmployeeName} was approved by Director {approvedByDirectorUserName}." +
                                      (!string.IsNullOrWhiteSpace(directorRemarks) ? $" Remarks: {directorRemarks}" : ""),
                            SenderName = approvedByDirectorUserName,
                            TargetUserId = request.OriginatorUserName,
                            ActionUrl = "/payroll/manual-attendance"
                        });
                    }
                    catch (Exception notifEx)
                    {
                        _logger.LogError(notifEx, "Failed to send approval notification to originator '{Originator}' for task #{TaskId}", request.OriginatorUserName, taskId);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error approving attendance task #{TaskId}", taskId);
                throw;
            }
        }

        public async Task<bool> RejectAttendanceActionAsync(int taskId, string rejectedByDirectorUserName, string rejectionReason)
        {
            try
            {
                var request = await GetAttendanceTaskDetailsAsync(taskId);
                if (request == null)
                {
                    throw new InvalidOperationException($"Attendance task #{taskId} details could not be found.");
                }

                string entityRefId = $"{request.ActionType}_{request.EmpID}_{request.Year:D4}{request.Month:D2}" +
                                     (request.AttendanceDate.HasValue ? $"_{request.AttendanceDate.Value:dd}" : "");

                await RejectRoleTaskAsync("AttendanceWorkflow", entityRefId, "Director", rejectedByDirectorUserName,
                    $"Attendance change request rejected by Director {rejectedByDirectorUserName}. Reason: {rejectionReason}");

                // Notify originator
                if (!string.IsNullOrWhiteSpace(request.OriginatorUserName) && !string.Equals(request.OriginatorUserName, "System", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        await _notificationService.SendNotificationAsync(new AppNotification
                        {
                            Category = NotificationCategory.Task,
                            Title = $"❌ Attendance Rejected: [{request.EmpID}] {request.EmployeeName}",
                            Message = $"Your attendance change request for [{request.EmpID}] {request.EmployeeName} was rejected by Director {rejectedByDirectorUserName}. Remarks: {rejectionReason}",
                            SenderName = rejectedByDirectorUserName,
                            TargetUserId = request.OriginatorUserName,
                            ActionUrl = "/payroll/manual-attendance"
                        });
                    }
                    catch (Exception notifEx)
                    {
                        _logger.LogError(notifEx, "Failed to send rejection notification to originator '{Originator}' for task #{TaskId}", request.OriginatorUserName, taskId);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error rejecting attendance task #{TaskId}", taskId);
                throw;
            }
        }

        public async Task<List<TaskItem>> GetPendingAttendanceTasksAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT t.*, 
                           t.Description
                    FROM TaskItems t
                    WHERE t.SourceEntityType = 'AttendanceWorkflow' 
                      AND t.Status IN (0, 1)
                    ORDER BY t.CreatedAt DESC";
                var items = await db.QueryAsync<TaskItem>(sql);
                return items.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching pending attendance tasks");
                return new List<TaskItem>();
            }
        }

        public async Task<AttendanceWorkflowRequestDto?> GetAttendanceTaskDetailsAsync(int taskId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT TOP 1 Content FROM TaskComments 
                    WHERE TaskId = @TaskId 
                    ORDER BY Id ASC";
                string? json = await db.ExecuteScalarAsync<string>(sql, new { TaskId = taskId });
                if (string.IsNullOrWhiteSpace(json)) return null;

                var dto = JsonSerializer.Deserialize<AttendanceWorkflowRequestDto>(json);
                if (dto != null)
                {
                    dto.TaskId = taskId;
                }
                return dto;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error parsing attendance task details for task #{TaskId}", taskId);
                return null;
            }
        }

        #endregion

        #region Maker Item Rate Approval Workflow

        public async Task<int> RequestMakerItemRateChangeAsync(MakerItemRateWorkflowRequestDto request)
        {
            try
            {
                string entityRefId = request.EntryID.ToString();

                using var db = CreateConnection();

                // Check for existing active request
                const string checkSql = @"
                    SELECT TOP 1 Id FROM TaskItems 
                    WHERE SourceEntityType = 'MakerItemRateWorkflow' 
                      AND SourceEntityRefId = @RefId 
                      AND Status IN (0, 1)";
                var existingTaskId = await db.ExecuteScalarAsync<int?>(checkSql, new { RefId = entityRefId });
                if (existingTaskId.HasValue && existingTaskId.Value > 0)
                {
                    _logger.LogInformation("Active maker rate workflow task #{TaskId} already exists for EntryID {EntryId}", existingTaskId.Value, entityRefId);
                    return existingTaskId.Value;
                }

                string title = $"💰 Rate Approval: [{request.ItemID}] {request.ItemName} - {request.MakerName}";
                string desc = $"Maker: {request.MakerName} | Process: {request.ProcessName} | Item: [{request.ItemID}] {request.ItemName} | Rate: {request.OldRate:N2} -> {request.NewRate:N2} | Reason: {request.Reason} | Requested by: {request.OriginatorUserName}";

                var taskId = await CreateRoleTaskAsync(new WorkflowTaskCreateRequest
                {
                    SourceEntityType = "MakerItemRateWorkflow",
                    SourceEntityRefId = entityRefId,
                    TargetRole = "Director",
                    Title = title,
                    Description = desc,
                    ActionUrl = $"/production/maker-rate-approval/{0}",
                    Priority = 2,
                    DueDate = DateTime.Today.AddDays(2),
                    CreatedBy = request.OriginatorUserName
                });

                // Update ActionUrl to point directly to this task
                const string updateUrlSql = "UPDATE TaskItems SET ActionUrl = @ActionUrl WHERE Id = @TaskId";
                await db.ExecuteAsync(updateUrlSql, new { ActionUrl = $"/production/maker-rate-approval/{taskId}", TaskId = taskId });

                request.TaskId = taskId;
                string payloadJson = JsonSerializer.Serialize(request);

                // Store serialized payload in TaskComments for lossless replay
                const string commentSql = @"
                    INSERT INTO TaskComments (TaskId, UserId, Content, CreatedAt)
                    VALUES (@TaskId, @UserId, @Content, GETUTCDATE())";
                await db.ExecuteAsync(commentSql, new
                {
                    TaskId = taskId,
                    UserId = request.OriginatorUserName,
                    Content = payloadJson
                });

                return taskId;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error requesting maker rate workflow task for EntryID {EntryId}", request.EntryID);
                throw;
            }
        }

        public async Task<bool> ApproveMakerItemRateActionAsync(int taskId, string approvedByDirectorUserName, string? directorRemarks = null)
        {
            try
            {
                var request = await GetMakerRateTaskDetailsAsync(taskId);
                if (request == null)
                {
                    throw new InvalidOperationException($"Maker rate task #{taskId} details could not be found.");
                }

                // Update assigned item rate in DB
                await _makerItemData.UpdateAssignedItemRateAndRemarksAsync(
                    request.EntryID,
                    request.NewRate,
                    request.OldRate,
                    request.Remarks,
                    approvedByDirectorUserName);

                string entityRefId = request.EntryID.ToString();

                // Mark task complete
                await CompleteRoleTaskAsync("MakerItemRateWorkflow", entityRefId, "Director", approvedByDirectorUserName,
                    $"Rate change approved and applied by Director {approvedByDirectorUserName}." +
                    (!string.IsNullOrWhiteSpace(directorRemarks) ? $" Remarks: {directorRemarks}" : ""));

                // Notify originator
                if (!string.IsNullOrWhiteSpace(request.OriginatorUserName) && !string.Equals(request.OriginatorUserName, "System", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        await _notificationService.SendNotificationAsync(new AppNotification
                        {
                            Category = NotificationCategory.Task,
                            Title = $"✅ Rate Approved: [{request.ItemID}] {request.ItemName}",
                            Message = $"Your rate change request for [{request.ItemID}] ({request.MakerName}) from {request.OldRate:N2} to {request.NewRate:N2} was approved by Director {approvedByDirectorUserName}." +
                                      (!string.IsNullOrWhiteSpace(directorRemarks) ? $" Remarks: {directorRemarks}" : ""),
                            SenderName = approvedByDirectorUserName,
                            TargetUserId = request.OriginatorUserName,
                            ActionUrl = $"/production/maker-item-assignment?makerId={request.VendID}&processId={request.ProcessID}"
                        });
                    }
                    catch (Exception notifEx)
                    {
                        _logger.LogError(notifEx, "Failed to send rate approval notification to originator '{Originator}' for task #{TaskId}", request.OriginatorUserName, taskId);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error approving maker rate task #{TaskId}", taskId);
                throw;
            }
        }

        public async Task<bool> RejectMakerItemRateActionAsync(int taskId, string rejectedByDirectorUserName, string rejectionReason)
        {
            try
            {
                var request = await GetMakerRateTaskDetailsAsync(taskId);
                if (request == null)
                {
                    throw new InvalidOperationException($"Maker rate task #{taskId} details could not be found.");
                }

                string entityRefId = request.EntryID.ToString();

                await RejectRoleTaskAsync("MakerItemRateWorkflow", entityRefId, "Director", rejectedByDirectorUserName,
                    $"Rate change request rejected by Director {rejectedByDirectorUserName}. Reason: {rejectionReason}");

                // Notify originator
                if (!string.IsNullOrWhiteSpace(request.OriginatorUserName) && !string.Equals(request.OriginatorUserName, "System", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        await _notificationService.SendNotificationAsync(new AppNotification
                        {
                            Category = NotificationCategory.Task,
                            Title = $"❌ Rate Rejected: [{request.ItemID}] {request.ItemName}",
                            Message = $"Your rate change request for [{request.ItemID}] ({request.MakerName}) was rejected by Director {rejectedByDirectorUserName}. Reason: {rejectionReason}",
                            SenderName = rejectedByDirectorUserName,
                            TargetUserId = request.OriginatorUserName,
                            ActionUrl = $"/production/maker-item-assignment?makerId={request.VendID}&processId={request.ProcessID}"
                        });
                    }
                    catch (Exception notifEx)
                    {
                        _logger.LogError(notifEx, "Failed to send rate rejection notification to originator '{Originator}' for task #{TaskId}", request.OriginatorUserName, taskId);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error rejecting maker rate task #{TaskId}", taskId);
                throw;
            }
        }

        public async Task<List<TaskItem>> GetPendingMakerRateTasksAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT t.*, 
                           t.Description
                    FROM TaskItems t
                    WHERE t.SourceEntityType = 'MakerItemRateWorkflow' 
                      AND t.Status IN (0, 1)
                    ORDER BY t.CreatedAt DESC";
                var items = await db.QueryAsync<TaskItem>(sql);
                return items.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching pending maker rate tasks");
                return new List<TaskItem>();
            }
        }

        public async Task<MakerItemRateWorkflowRequestDto?> GetMakerRateTaskDetailsAsync(int taskId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT TOP 1 Content FROM TaskComments 
                    WHERE TaskId = @TaskId 
                    ORDER BY Id ASC";
                string? json = await db.ExecuteScalarAsync<string>(sql, new { TaskId = taskId });
                if (string.IsNullOrWhiteSpace(json)) return null;

                var dto = JsonSerializer.Deserialize<MakerItemRateWorkflowRequestDto>(json);
                if (dto != null)
                {
                    dto.TaskId = taskId;
                }
                return dto;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error parsing maker rate task details for task #{TaskId}", taskId);
                return null;
            }
        }

        #endregion

        #region Production Deletion Requests Workflow

        public async Task<int> RequestLotReceivingDeletionAsync(ProductionDeletionRequestModel request)
        {
            try
            {
                // 1. Create request record in ProductionDeletionRequests
                int requestId = await _prodDeletionData.CreateRequestAsync(request);
                request.Id = requestId;

                // 2. Create TaskItems record for Director
                string title = $"⚠️ Lot Receiving Deletion: Lot #{request.LotNo} ({request.ItemName})";
                string desc = $"Lot: #{request.LotNo} | Process: {request.ProcessName} | Item: [{request.ItemCode}] {request.ItemName} | Qty: {request.Qty:N2} | Maker: {request.MakerName} | Reason: {request.Reason} | Requested by: {request.RequestedBy}";

                var taskId = await CreateRoleTaskAsync(new WorkflowTaskCreateRequest
                {
                    SourceEntityType = "DeleteLotAuthorization",
                    SourceEntityRefId = requestId.ToString(),
                    TargetRole = "Director",
                    Title = title,
                    Description = desc,
                    ActionUrl = $"/production/lot-deletion-approval/{0}",
                    Priority = 3, // Urgent
                    DueDate = DateTime.Today.AddDays(1),
                    CreatedBy = request.RequestedBy
                });

                using var db = CreateConnection();
                const string updateUrlSql = "UPDATE TaskItems SET ActionUrl = @ActionUrl WHERE Id = @TaskId";
                await db.ExecuteAsync(updateUrlSql, new { ActionUrl = $"/production/lot-deletion-approval/{taskId}", TaskId = taskId });

                // Link TaskId back to ProductionDeletionRequests
                await _prodDeletionData.UpdateTaskIdAsync(requestId, taskId);

                // Add comment for audit
                const string commentSql = @"
                    INSERT INTO TaskComments (TaskId, UserId, Content, CreatedAt)
                    VALUES (@TaskId, @UserId, @Content, GETUTCDATE())";
                await db.ExecuteAsync(commentSql, new
                {
                    TaskId = taskId,
                    UserId = request.RequestedBy,
                    Content = $"[Deletion Request #{requestId}] {request.Reason}" + (!string.IsNullOrWhiteSpace(request.MachineName) ? $" (Station: {request.MachineName})" : "")
                });

                return requestId;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error requesting lot receiving deletion for Lot #{LotNo}, VRD_EntryID {EntityRefId}", request.LotNo, request.EntityRefID);
                throw;
            }
        }

        public async Task<bool> ApproveLotReceivingDeletionAsync(int requestId, string approvedByDirectorUserName, string? directorRemarks = null)
        {
            try
            {
                var request = await _prodDeletionData.GetRequestByIdAsync(requestId);
                if (request == null)
                    throw new InvalidOperationException($"Production deletion request #{requestId} not found.");

                if (request.Status != "Pending")
                    throw new InvalidOperationException($"Production deletion request #{requestId} is already {request.Status}.");

                // Final check: confirm no linked downstream issuances
                int issCount = await _makerRcvListData.CheckIssuanceExistsAsync(request.EntityRefID);
                if (issCount > 0)
                {
                    throw new InvalidOperationException($"Cannot delete Lot #{request.LotNo}: {issCount} linked downstream issuance(s) found. Delete them first or reject this request.");
                }

                // 1. Delete receiving record (archives to VendRcvdDetail_Deletions & deletes from VendRcvdDetail)
                bool deleted = await _makerRcvListData.DeleteReceivingAsync(request.EntityRefID, approvedByDirectorUserName, Environment.MachineName);
                if (!deleted)
                {
                    throw new InvalidOperationException($"Failed to delete receiving record for Lot #{request.LotNo}.");
                }

                // 2. Mark ProductionDeletionRequests approved
                await _prodDeletionData.ApproveRequestAsync(requestId, approvedByDirectorUserName, directorRemarks);

                // 3. Mark TaskItem completed
                if (request.TaskId.HasValue && request.TaskId.Value > 0)
                {
                    using var db = CreateConnection();
                    const string completeSql = @"
                        UPDATE TaskItems 
                        SET Status = 2, 
                            CompletedAt = GETUTCDATE(), 
                            CompletedBy = @ApprovedBy, 
                            UpdatedAt = GETUTCDATE()
                        WHERE Id = @TaskId";
                    await db.ExecuteAsync(completeSql, new { ApprovedBy = approvedByDirectorUserName, TaskId = request.TaskId.Value });

                    if (!string.IsNullOrWhiteSpace(directorRemarks))
                    {
                        const string commentSql = @"
                            INSERT INTO TaskComments (TaskId, UserId, Content, CreatedAt)
                            VALUES (@TaskId, @UserId, @Content, GETUTCDATE())";
                        await db.ExecuteAsync(commentSql, new
                        {
                            TaskId = request.TaskId.Value,
                            UserId = approvedByDirectorUserName,
                            Content = $"[Director Approved & Deleted] {directorRemarks}"
                        });
                    }

                    // Notify originator
                    if (!string.IsNullOrWhiteSpace(request.RequestedBy) && !string.Equals(request.RequestedBy, "System", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            await _notificationService.SendNotificationAsync(new AppNotification
                            {
                                Category = NotificationCategory.Task,
                                Title = $"✅ Lot #{request.LotNo} Deletion Approved",
                                Message = $"Your deletion request for Lot #{request.LotNo} has been approved and deleted by Director {approvedByDirectorUserName}." +
                                          (!string.IsNullOrWhiteSpace(directorRemarks) ? $" Remarks: {directorRemarks}" : ""),
                                SenderName = approvedByDirectorUserName,
                                TargetUserId = request.RequestedBy,
                                ActionUrl = "/production/receiving-list"
                            });
                        }
                        catch (Exception notifEx)
                        {
                            _logger.LogError(notifEx, "Failed to send deletion approval notification to '{Originator}'", request.RequestedBy);
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error approving lot deletion request #{RequestId}", requestId);
                throw;
            }
        }

        public async Task<bool> RejectLotReceivingDeletionAsync(int requestId, string rejectedByDirectorUserName, string rejectionReason)
        {
            try
            {
                var request = await _prodDeletionData.GetRequestByIdAsync(requestId);
                if (request == null)
                    throw new InvalidOperationException($"Production deletion request #{requestId} not found.");

                if (request.Status != "Pending")
                    throw new InvalidOperationException($"Production deletion request #{requestId} is already {request.Status}.");

                // 1. Mark ProductionDeletionRequests rejected
                await _prodDeletionData.RejectRequestAsync(requestId, rejectedByDirectorUserName, rejectionReason);

                // 2. Mark TaskItem rejected
                if (request.TaskId.HasValue && request.TaskId.Value > 0)
                {
                    using var db = CreateConnection();
                    const string rejectSql = @"
                        UPDATE TaskItems 
                        SET Status = 3, 
                            CompletedBy = @RejectedBy, 
                            UpdatedAt = GETUTCDATE()
                        WHERE Id = @TaskId";
                    await db.ExecuteAsync(rejectSql, new { RejectedBy = rejectedByDirectorUserName, TaskId = request.TaskId.Value });

                    const string commentSql = @"
                        INSERT INTO TaskComments (TaskId, UserId, Content, CreatedAt)
                        VALUES (@TaskId, @UserId, @Content, GETUTCDATE())";
                    await db.ExecuteAsync(commentSql, new
                    {
                        TaskId = request.TaskId.Value,
                        UserId = rejectedByDirectorUserName,
                        Content = $"[Director Rejected] {rejectionReason}"
                    });

                    // Notify originator
                    if (!string.IsNullOrWhiteSpace(request.RequestedBy) && !string.Equals(request.RequestedBy, "System", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            await _notificationService.SendNotificationAsync(new AppNotification
                            {
                                Category = NotificationCategory.Task,
                                Title = $"❌ Lot #{request.LotNo} Deletion Rejected",
                                Message = $"Your deletion request for Lot #{request.LotNo} was rejected by Director {rejectedByDirectorUserName}. Reason: {rejectionReason}",
                                SenderName = rejectedByDirectorUserName,
                                TargetUserId = request.RequestedBy,
                                ActionUrl = "/production/receiving-list"
                            });
                        }
                        catch (Exception notifEx)
                        {
                            _logger.LogError(notifEx, "Failed to send deletion rejection notification to '{Originator}'", request.RequestedBy);
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error rejecting lot deletion request #{RequestId}", requestId);
                throw;
            }
        }

        public Task<List<ProductionDeletionRequestModel>> GetPendingLotDeletionRequestsAsync()
        {
            return _prodDeletionData.GetPendingRequestsAsync("LotReceiving");
        }

        public Task<ProductionDeletionRequestModel?> GetLotDeletionRequestDetailsAsync(int requestId)
        {
            return _prodDeletionData.GetRequestByIdAsync(requestId);
        }

        public async Task<int> RequestProductionIssuanceDeletionAsync(ProductionDeletionRequestModel request)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(request.RequestType))
                {
                    request.RequestType = "LotIssuance";
                }

                // 1. Create request record in ProductionDeletionRequests
                int requestId = await _prodDeletionData.CreateRequestAsync(request);
                request.Id = requestId;

                // 2. Create TaskItems record for Director
                string title = $"⚠️ Issuance Deletion: PO #{request.OrderNo ?? request.LotNo} ({request.MakerName})";
                string desc = $"Order/Lot: #{request.OrderNo ?? request.LotNo} | Process: {request.ProcessName} | Item: [{request.ItemCode}] {request.ItemName} | Qty: {request.Qty:N2} | Maker: {request.MakerName} | Reason: {request.Reason} | Requested by: {request.RequestedBy}";

                var taskId = await CreateRoleTaskAsync(new WorkflowTaskCreateRequest
                {
                    SourceEntityType = "DeleteProductionIssuance",
                    SourceEntityRefId = requestId.ToString(),
                    TargetRole = "Director",
                    Title = title,
                    Description = desc,
                    ActionUrl = $"/production/issuance-deletion-approval/{0}",
                    Priority = 3, // Urgent
                    DueDate = DateTime.Today.AddDays(1),
                    CreatedBy = request.RequestedBy
                });

                using var db = CreateConnection();
                const string updateUrlSql = "UPDATE TaskItems SET ActionUrl = @ActionUrl WHERE Id = @TaskId";
                await db.ExecuteAsync(updateUrlSql, new { ActionUrl = $"/production/issuance-deletion-approval/{taskId}", TaskId = taskId });

                // Link TaskId back to ProductionDeletionRequests
                await _prodDeletionData.UpdateTaskIdAsync(requestId, taskId);

                // Add comment for audit
                const string commentSql = @"
                    INSERT INTO TaskComments (TaskId, UserId, Content, CreatedAt)
                    VALUES (@TaskId, @UserId, @Content, GETUTCDATE())";
                await db.ExecuteAsync(commentSql, new
                {
                    TaskId = taskId,
                    UserId = request.RequestedBy,
                    Content = $"[Issuance Deletion Request #{requestId}] {request.Reason}" + (!string.IsNullOrWhiteSpace(request.MachineName) ? $" (Station: {request.MachineName})" : "")
                });

                return requestId;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error requesting issuance deletion for Order/Lot #{OrderNo}, EntryID {EntityRefId}", request.OrderNo, request.EntityRefID);
                throw;
            }
        }

        public async Task<bool> ApproveProductionIssuanceDeletionAsync(int requestId, string approvedByDirectorUserName, string? directorRemarks = null)
        {
            try
            {
                var request = await _prodDeletionData.GetRequestByIdAsync(requestId);
                if (request == null)
                    throw new InvalidOperationException($"Production issuance deletion request #{requestId} not found.");

                if (request.Status != "Pending")
                    throw new InvalidOperationException($"Production issuance deletion request #{requestId} is already {request.Status}.");

                // Final check 1: confirm no goods received against this issuance
                int rcvCount = await _makerPOListData.CheckReceivingExistsAsync(request.EntityRefID);
                if (rcvCount > 0)
                {
                    throw new InvalidOperationException($"Cannot delete Issuance #{request.OrderNo ?? request.LotNo}: Receiving records already exist against this issuance. Delete receiving records first or reject this request.");
                }

                // Final check 2: check if short/long advance loan exists
                if (!string.IsNullOrWhiteSpace(request.OrderNo))
                {
                    var (shortLoan, longLoan) = await _makerPOListData.CheckLoanExistsAsync(request.OrderNo);
                    if (shortLoan || longLoan)
                    {
                        throw new InvalidOperationException($"Cannot delete Issuance #{request.OrderNo}: Short/Long term loan is linked to this order.");
                    }
                }

                // 1. Delete issuance record (reverts stock/SF quantities and deletes from VendIssued)
                bool deleted = await _makerPOListData.DeleteIssuanceAsync(request.EntityRefID);
                if (!deleted)
                {
                    throw new InvalidOperationException($"Failed to delete issuance record for EntryID {request.EntityRefID}.");
                }

                // 2. Mark ProductionDeletionRequests approved
                await _prodDeletionData.ApproveRequestAsync(requestId, approvedByDirectorUserName, directorRemarks);

                // 3. Mark TaskItem completed
                if (request.TaskId.HasValue && request.TaskId.Value > 0)
                {
                    using var db = CreateConnection();
                    const string completeSql = @"
                        UPDATE TaskItems 
                        SET Status = 2, 
                            CompletedAt = GETUTCDATE(), 
                            CompletedBy = @ApprovedBy, 
                            UpdatedAt = GETUTCDATE()
                        WHERE Id = @TaskId";
                    await db.ExecuteAsync(completeSql, new { ApprovedBy = approvedByDirectorUserName, TaskId = request.TaskId.Value });

                    if (!string.IsNullOrWhiteSpace(directorRemarks))
                    {
                        const string commentSql = @"
                            INSERT INTO TaskComments (TaskId, UserId, Content, CreatedAt)
                            VALUES (@TaskId, @UserId, @Content, GETUTCDATE())";
                        await db.ExecuteAsync(commentSql, new
                        {
                            TaskId = request.TaskId.Value,
                            UserId = approvedByDirectorUserName,
                            Content = $"[Director Approved & Deleted] {directorRemarks}"
                        });
                    }

                    // Notify originator
                    if (!string.IsNullOrWhiteSpace(request.RequestedBy) && !string.Equals(request.RequestedBy, "System", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            await _notificationService.SendNotificationAsync(new AppNotification
                            {
                                Category = NotificationCategory.Task,
                                Title = $"✅ Issuance Deletion Approved",
                                Message = $"Your deletion request for Issuance #{request.OrderNo ?? request.LotNo} has been approved and deleted by Director {approvedByDirectorUserName}." +
                                          (!string.IsNullOrWhiteSpace(directorRemarks) ? $" Remarks: {directorRemarks}" : ""),
                                SenderName = approvedByDirectorUserName,
                                TargetUserId = request.RequestedBy,
                                ActionUrl = "/production/maker-po-list"
                            });
                        }
                        catch (Exception notifEx)
                        {
                            _logger.LogError(notifEx, "Failed to send issuance deletion approval notification to '{Originator}'", request.RequestedBy);
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error approving issuance deletion request #{RequestId}", requestId);
                throw;
            }
        }

        public async Task<bool> RejectProductionIssuanceDeletionAsync(int requestId, string rejectedByDirectorUserName, string rejectionReason)
        {
            try
            {
                var request = await _prodDeletionData.GetRequestByIdAsync(requestId);
                if (request == null)
                    throw new InvalidOperationException($"Production issuance deletion request #{requestId} not found.");

                if (request.Status != "Pending")
                    throw new InvalidOperationException($"Production issuance deletion request #{requestId} is already {request.Status}.");

                // 1. Mark ProductionDeletionRequests rejected
                await _prodDeletionData.RejectRequestAsync(requestId, rejectedByDirectorUserName, rejectionReason);

                // 2. Mark TaskItem rejected
                if (request.TaskId.HasValue && request.TaskId.Value > 0)
                {
                    using var db = CreateConnection();
                    const string rejectSql = @"
                        UPDATE TaskItems 
                        SET Status = 3, 
                            CompletedBy = @RejectedBy, 
                            UpdatedAt = GETUTCDATE()
                        WHERE Id = @TaskId";
                    await db.ExecuteAsync(rejectSql, new { RejectedBy = rejectedByDirectorUserName, TaskId = request.TaskId.Value });

                    const string commentSql = @"
                        INSERT INTO TaskComments (TaskId, UserId, Content, CreatedAt)
                        VALUES (@TaskId, @UserId, @Content, GETUTCDATE())";
                    await db.ExecuteAsync(commentSql, new
                    {
                        TaskId = request.TaskId.Value,
                        UserId = rejectedByDirectorUserName,
                        Content = $"[Director Rejected] {rejectionReason}"
                    });

                    // Notify originator
                    if (!string.IsNullOrWhiteSpace(request.RequestedBy) && !string.Equals(request.RequestedBy, "System", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            await _notificationService.SendNotificationAsync(new AppNotification
                            {
                                Category = NotificationCategory.Task,
                                Title = $"❌ Issuance Deletion Rejected",
                                Message = $"Your deletion request for Issuance #{request.OrderNo ?? request.LotNo} was rejected by Director {rejectedByDirectorUserName}. Reason: {rejectionReason}",
                                SenderName = rejectedByDirectorUserName,
                                TargetUserId = request.RequestedBy,
                                ActionUrl = "/production/maker-po-list"
                            });
                        }
                        catch (Exception notifEx)
                        {
                            _logger.LogError(notifEx, "Failed to send issuance deletion rejection notification to '{Originator}'", request.RequestedBy);
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error rejecting issuance deletion request #{RequestId}", requestId);
                throw;
            }
        }

        public Task<List<ProductionDeletionRequestModel>> GetPendingIssuanceDeletionRequestsAsync()
        {
            return _prodDeletionData.GetPendingRequestsAsync("LotIssuance");
        }

        public async Task<int> RequestSkipProcessAsync(ProductionDeletionRequestModel request)
        {
            try
            {
                request.RequestType = "SkipProcess";

                // 1. Create request record in ProductionDeletionRequests
                int requestId = await _prodDeletionData.CreateRequestAsync(request);
                request.Id = requestId;

                // 2. Create TaskItems record for Director
                string title = $"⚠️ Skip Process: Lot #{request.LotNo} -> {request.TargetProcessName}";
                string desc = $"Lot: #{request.LotNo} | Item: [{request.ItemCode}] {request.ItemName} | Current Process: {request.ProcessName} | Target Process: {request.TargetProcessName} | Qty: {request.Qty:N0} | Reason: {request.Reason} | Requested by: {request.RequestedBy}";

                var taskId = await CreateRoleTaskAsync(new WorkflowTaskCreateRequest
                {
                    SourceEntityType = "SkipProcess",
                    SourceEntityRefId = requestId.ToString(),
                    TargetRole = "Director",
                    Title = title,
                    Description = desc,
                    ActionUrl = $"/production/skip-process-approval/{0}",
                    Priority = 3, // Urgent
                    DueDate = DateTime.Today.AddDays(1),
                    CreatedBy = request.RequestedBy
                });

                using var db = CreateConnection();
                const string updateUrlSql = "UPDATE TaskItems SET ActionUrl = @ActionUrl WHERE Id = @TaskId";
                await db.ExecuteAsync(updateUrlSql, new { ActionUrl = $"/production/skip-process-approval/{taskId}", TaskId = taskId });

                // Link TaskId back to ProductionDeletionRequests
                await _prodDeletionData.UpdateTaskIdAsync(requestId, taskId);

                // Add comment for audit
                const string commentSql = @"
                    INSERT INTO TaskComments (TaskId, UserId, Content, CreatedAt)
                    VALUES (@TaskId, @UserId, @Content, GETUTCDATE())";
                await db.ExecuteAsync(commentSql, new
                {
                    TaskId = taskId,
                    UserId = request.RequestedBy,
                    Content = $"[Skip Process Request #{requestId}] {request.Reason} (From: {request.ProcessName} -> To: {request.TargetProcessName})"
                });

                return requestId;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error requesting skip process for Lot #{LotNo}", request.LotNo);
                throw;
            }
        }

        public async Task<bool> ApproveSkipProcessAsync(int requestId, string approvedByDirectorUserName, string? directorRemarks = null)
        {
            try
            {
                var request = await _prodDeletionData.GetRequestByIdAsync(requestId);
                if (request == null)
                    throw new InvalidOperationException($"Skip process request #{requestId} not found.");

                if (request.Status != "Pending")
                    throw new InvalidOperationException($"Skip process request #{requestId} is already {request.Status}.");

                if (string.IsNullOrWhiteSpace(request.ItemCode) || !request.ProcessID.HasValue || request.TargetProcessID <= 0 || string.IsNullOrWhiteSpace(request.LotNo))
                {
                    throw new InvalidOperationException($"Skip process request #{requestId} contains invalid parameters.");
                }

                // 1. Execute Skip Process (updates NextProcessID in VendRcvdDetail)
                bool skipped = await _lotIssuanceData.SkipProcessAsync(
                    request.ItemCode,
                    request.ProcessID.Value,
                    request.TargetProcessID,
                    request.LotNo);

                if (!skipped)
                {
                    throw new InvalidOperationException($"Failed to skip process for Lot #{request.LotNo}. Please ensure the lot is available at the specified process.");
                }

                // 2. Mark ProductionDeletionRequests approved
                await _prodDeletionData.ApproveRequestAsync(requestId, approvedByDirectorUserName, directorRemarks);

                // 3. Mark TaskItem completed
                if (request.TaskId.HasValue && request.TaskId.Value > 0)
                {
                    using var db = CreateConnection();
                    const string completeSql = @"
                        UPDATE TaskItems 
                        SET Status = 2, 
                            CompletedAt = GETUTCDATE(), 
                            CompletedBy = @ApprovedBy, 
                            UpdatedAt = GETUTCDATE()
                        WHERE Id = @TaskId";
                    await db.ExecuteAsync(completeSql, new { ApprovedBy = approvedByDirectorUserName, TaskId = request.TaskId.Value });

                    if (!string.IsNullOrWhiteSpace(directorRemarks))
                    {
                        const string commentSql = @"
                            INSERT INTO TaskComments (TaskId, UserId, Content, CreatedAt)
                            VALUES (@TaskId, @UserId, @Content, GETUTCDATE())";
                        await db.ExecuteAsync(commentSql, new
                        {
                            TaskId = request.TaskId.Value,
                            UserId = approvedByDirectorUserName,
                            Content = $"[Director Approved & Skipped] {directorRemarks}"
                        });
                    }

                    // Notify originator
                    if (!string.IsNullOrWhiteSpace(request.RequestedBy) && !string.Equals(request.RequestedBy, "System", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            await _notificationService.SendNotificationAsync(new AppNotification
                            {
                                Category = NotificationCategory.Task,
                                Title = $"✅ Skip Process Approved: Lot #{request.LotNo}",
                                Message = $"Your request to skip Lot #{request.LotNo} to '{request.TargetProcessName}' was approved by Director {approvedByDirectorUserName}." +
                                          (!string.IsNullOrWhiteSpace(directorRemarks) ? $" Remarks: {directorRemarks}" : ""),
                                SenderName = approvedByDirectorUserName,
                                TargetUserId = request.RequestedBy,
                                ActionUrl = "/production/lot-issuance"
                            });
                        }
                        catch (Exception notifEx)
                        {
                            _logger.LogError(notifEx, "Failed to send skip process approval notification to '{Originator}'", request.RequestedBy);
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error approving skip process request #{RequestId}", requestId);
                throw;
            }
        }

        public async Task<bool> RejectSkipProcessAsync(int requestId, string rejectedByDirectorUserName, string rejectionReason)
        {
            try
            {
                var request = await _prodDeletionData.GetRequestByIdAsync(requestId);
                if (request == null)
                    throw new InvalidOperationException($"Skip process request #{requestId} not found.");

                if (request.Status != "Pending")
                    throw new InvalidOperationException($"Skip process request #{requestId} is already {request.Status}.");

                // 1. Mark ProductionDeletionRequests rejected
                await _prodDeletionData.RejectRequestAsync(requestId, rejectedByDirectorUserName, rejectionReason);

                // 2. Mark TaskItem rejected
                if (request.TaskId.HasValue && request.TaskId.Value > 0)
                {
                    using var db = CreateConnection();
                    const string rejectSql = @"
                        UPDATE TaskItems 
                        SET Status = 3, 
                            CompletedBy = @RejectedBy, 
                            UpdatedAt = GETUTCDATE()
                        WHERE Id = @TaskId";
                    await db.ExecuteAsync(rejectSql, new { RejectedBy = rejectedByDirectorUserName, TaskId = request.TaskId.Value });

                    const string commentSql = @"
                        INSERT INTO TaskComments (TaskId, UserId, Content, CreatedAt)
                        VALUES (@TaskId, @UserId, @Content, GETUTCDATE())";
                    await db.ExecuteAsync(commentSql, new
                    {
                        TaskId = request.TaskId.Value,
                        UserId = rejectedByDirectorUserName,
                        Content = $"[Director Rejected] {rejectionReason}"
                    });

                    // Notify originator
                    if (!string.IsNullOrWhiteSpace(request.RequestedBy) && !string.Equals(request.RequestedBy, "System", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            await _notificationService.SendNotificationAsync(new AppNotification
                            {
                                Category = NotificationCategory.Task,
                                Title = $"❌ Skip Process Rejected: Lot #{request.LotNo}",
                                Message = $"Your request to skip Lot #{request.LotNo} to '{request.TargetProcessName}' was rejected by Director {rejectedByDirectorUserName}. Reason: {rejectionReason}",
                                SenderName = rejectedByDirectorUserName,
                                TargetUserId = request.RequestedBy,
                                ActionUrl = "/production/lot-issuance"
                            });
                        }
                        catch (Exception notifEx)
                        {
                            _logger.LogError(notifEx, "Failed to send skip process rejection notification to '{Originator}'", request.RequestedBy);
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error rejecting skip process request #{RequestId}", requestId);
                throw;
            }
        }

        public Task<List<ProductionDeletionRequestModel>> GetPendingSkipProcessRequestsAsync()
        {
            return _prodDeletionData.GetPendingRequestsAsync("SkipProcess");
        }

        #endregion

        #region Governance Workflow Dynamic Checks

        public async Task<bool> IsApprovalRequiredAsync(string workflowCode, string userName, System.Security.Claims.ClaimsPrincipal? claimsPrincipal = null)
        {
            try
            {
                var roles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                if (string.Equals(userName, "admin", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(userName, "administrator", StringComparison.OrdinalIgnoreCase))
                {
                    roles.Add("Admin");
                    roles.Add("Administrator");
                    roles.Add("Director");
                }

                if (claimsPrincipal != null)
                {
                    foreach (var claim in claimsPrincipal.Claims.Where(c => c.Type == System.Security.Claims.ClaimTypes.Role))
                    {
                        roles.Add(claim.Value);
                    }
                }

                var dbUser = await _userData.GetUserByUserNameAsync(userName);
                if (dbUser != null)
                {
                    var userDbRoles = await _userRoleData.GetRolesByUserIdAsync(dbUser.UserID);
                    if (userDbRoles != null)
                    {
                        foreach (var r in userDbRoles) roles.Add(r);
                    }
                }

                return await _workflowConfigData.IsWorkflowApprovalRequiredAsync(workflowCode, roles);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error checking approval requirement for workflow {WorkflowCode}, user {UserName}", workflowCode, userName);
                return true;
            }
        }

        public Task<DataAccessLibrary.Models.Setup.WorkflowConfigurationModel?> GetWorkflowConfigAsync(string workflowCode)
        {
            return _workflowConfigData.GetWorkflowConfigByCodeAsync(workflowCode);
        }

        #endregion
    }
}
