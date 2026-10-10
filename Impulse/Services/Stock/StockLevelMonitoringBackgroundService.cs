using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.Stock
{
    public class StockLevelMonitoringBackgroundService : BackgroundService
    {
        private readonly IServiceScopeFactory _scopeFactory;
        private readonly ILogger<StockLevelMonitoringBackgroundService> _logger;
        private static readonly TimeSpan IntervalCheck = TimeSpan.FromMinutes(1);
        private DateTime _lastRunDate = DateTime.MinValue;

        public StockLevelMonitoringBackgroundService(
            IServiceScopeFactory scopeFactory,
            ILogger<StockLevelMonitoringBackgroundService> logger)
        {
            _scopeFactory = scopeFactory;
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _logger.LogInformation("StockLevelMonitoringBackgroundService has started. Scheduled for daily 08:00 AM digest.");

            // Wait 20 seconds after app startup
            try
            {
                await Task.Delay(TimeSpan.FromSeconds(20), stoppingToken);
            }
            catch (OperationCanceledException)
            {
                return;
            }

            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    var now = DateTime.Now;
                    // Trigger daily at 8:00 AM (or next check between 8:00 AM and 8:15 AM if app was started)
                    if (now.Hour == 8 && now.Date > _lastRunDate.Date)
                    {
                        _lastRunDate = now.Date;
                        _logger.LogInformation("StockLevelMonitoringBackgroundService: Executing 08:00 AM daily stock digest.");

                        using var scope = _scopeFactory.CreateScope();
                        var stockNotifService = scope.ServiceProvider.GetRequiredService<IStockNotificationService>();
                        await stockNotifService.SendDailyStockDigestAsync();
                    }
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Unexpected error in StockLevelMonitoringBackgroundService execution cycle.");
                }

                try
                {
                    await Task.Delay(IntervalCheck, stoppingToken);
                }
                catch (OperationCanceledException)
                {
                    break;
                }
            }
        }
    }
}
