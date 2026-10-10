using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Setup;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using Impulse.Services.IntraOffice;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.Stock
{
    public class StockNotificationService : IStockNotificationService
    {
        private readonly IRMDataAccess _rmDataAccess;
        private readonly IUserRoleDataAccess _userRoleDataAccess;
        private readonly IAppNotificationService _notificationService;
        private readonly ILogger<StockNotificationService> _logger;

        private const string StockRole = "Stock";

        public StockNotificationService(
            IRMDataAccess rmDataAccess,
            IUserRoleDataAccess userRoleDataAccess,
            IAppNotificationService notificationService,
            ILogger<StockNotificationService> logger)
        {
            _rmDataAccess = rmDataAccess;
            _userRoleDataAccess = userRoleDataAccess;
            _notificationService = notificationService;
            _logger = logger;
        }

        public async Task CheckAndNotifyStockMovementAsync(string materialCode, string transactionSource, string? refDocNo = null)
        {
            if (string.IsNullOrWhiteSpace(materialCode)) return;

            try
            {
                var itemStatus = await _rmDataAccess.GetMaterialStockStatusByCodeAsync(materialCode.Trim());
                if (itemStatus == null || itemStatus.BreachType == "Normal")
                {
                    return;
                }

                await DispatchSingleBreachAlertAsync(itemStatus, transactionSource, refDocNo);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to evaluate or dispatch stock notification for material {MaterialCode} from {Source}", materialCode, transactionSource);
            }
        }

        public async Task CheckAndNotifyStockMovementByIdAsync(int rmId, string transactionSource, string? refDocNo = null)
        {
            if (rmId <= 0) return;

            try
            {
                var itemStatus = await _rmDataAccess.GetMaterialStockStatusByIdAsync(rmId);
                if (itemStatus == null || itemStatus.BreachType == "Normal")
                {
                    return;
                }

                await DispatchSingleBreachAlertAsync(itemStatus, transactionSource, refDocNo);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to evaluate or dispatch stock notification for RMID {RMID} from {Source}", rmId, transactionSource);
            }
        }

        public async Task SendDailyStockDigestAsync()
        {
            try
            {
                var summary = await _rmDataAccess.GetStockBreachSummaryAsync();
                if (summary == null || summary.TotalBreaches == 0)
                {
                    _logger.LogInformation("Daily stock digest: All stock levels are within configured thresholds. No notifications needed.");
                    return;
                }

                var targetUsers = await GetStockRoleUserNamesAsync();

                var title = "Daily Stock Level Digest";
                var parts = new List<string>();

                if (summary.BelowMinCount > 0)
                    parts.Add($"{summary.BelowMinCount} below Min");
                if (summary.BelowReorderCount > 0)
                    parts.Add($"{summary.BelowReorderCount} at Reorder");
                if (summary.AboveMaxCount > 0)
                    parts.Add($"{summary.AboveMaxCount} Overstocked");

                var message = $"Morning Stock Summary: {string.Join(", ", parts)}. Click to view items requiring attention.";
                var actionUrl = "/rmlist?filter=breaches";

                if (targetUsers.Count == 0)
                {
                    // Broadcast notification if no specific users explicitly in Stock role
                    await _notificationService.SendNotificationAsync(new AppNotification
                    {
                        Category = NotificationCategory.Stock,
                        Title = title,
                        Message = message,
                        ActionUrl = actionUrl,
                        SenderName = "Stock Monitor"
                    });
                }
                else
                {
                    foreach (var userName in targetUsers)
                    {
                        await _notificationService.SendNotificationAsync(new AppNotification
                        {
                            Category = NotificationCategory.Stock,
                            Title = title,
                            Message = message,
                            ActionUrl = actionUrl,
                            TargetUserId = userName,
                            SenderName = "Stock Monitor"
                        });
                    }
                }

                _logger.LogInformation("Daily stock digest sent successfully for {Count} breach items to {UsersCount} recipients.", summary.TotalBreaches, targetUsers.Count);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error occurred while executing SendDailyStockDigestAsync.");
            }
        }

        private async Task DispatchSingleBreachAlertAsync(StockBreachItemDto item, string source, string? docNo)
        {
            var docSuffix = !string.IsNullOrWhiteSpace(docNo) ? $" (Ref: {docNo})" : "";
            string title;
            string message;
            string actionUrl;

            switch (item.BreachType)
            {
                case "MinStock":
                    title = $"Critical: Min Stock Alert - {item.RMID1}";
                    message = $"{item.RMName} current stock is {item.QtyInStock:N1} {item.Unit}, which is at or below the minimum safety level of {item.MinLevel:N1} {item.Unit} following {source}{docSuffix}.";
                    actionUrl = "/rmlist?filter=min";
                    break;

                case "Reorder":
                    title = $"Reorder Level Alert - {item.RMID1}";
                    message = $"{item.RMName} current stock is {item.QtyInStock:N1} {item.Unit}, reaching the reorder point of {item.ReorderPoint:N0} {item.Unit} following {source}{docSuffix}. Please initiate replenishment.";
                    actionUrl = "/rmlist?filter=reorder";
                    break;

                case "MaxStock":
                    title = $"Max Stock Alert - {item.RMID1}";
                    message = $"{item.RMName} current stock is {item.QtyInStock:N1} {item.Unit}, exceeding the maximum inventory capacity of {item.MaxLevel:N1} {item.Unit} following {source}{docSuffix}.";
                    actionUrl = "/rmlist?filter=max";
                    break;

                default:
                    return;
            }

            var targetUsers = await GetStockRoleUserNamesAsync();

            if (targetUsers.Count == 0)
            {
                await _notificationService.SendNotificationAsync(new AppNotification
                {
                    Category = NotificationCategory.Stock,
                    Title = title,
                    Message = message,
                    ActionUrl = actionUrl,
                    SenderName = "Stock Monitor"
                });
            }
            else
            {
                foreach (var user in targetUsers)
                {
                    await _notificationService.SendNotificationAsync(new AppNotification
                    {
                        Category = NotificationCategory.Stock,
                        Title = title,
                        Message = message,
                        ActionUrl = actionUrl,
                        TargetUserId = user,
                        SenderName = "Stock Monitor"
                    });
                }
            }
        }

        private async Task<HashSet<string>> GetStockRoleUserNamesAsync()
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var users = await _userRoleDataAccess.GetUsersByRoleAsync(StockRole);
                if (users != null)
                {
                    foreach (var u in users)
                    {
                        if (!u.InActive && !string.IsNullOrWhiteSpace(u.UserName))
                        {
                            set.Add(u.UserName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to resolve users for role '{Role}'", StockRole);
            }
            return set;
        }
    }
}
