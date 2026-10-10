using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public interface IStockNotificationService
    {
        /// <summary>
        /// Evaluates current stock against Min, Max, and Reorder thresholds for a specific material after a transaction
        /// and dispatches real-time alerts to the 'Stock' role if breached.
        /// </summary>
        Task CheckAndNotifyStockMovementAsync(string materialCode, string transactionSource, string? refDocNo = null);

        /// <summary>
        /// Evaluates current stock by RMID and dispatches real-time alerts to the 'Stock' role if breached.
        /// </summary>
        Task CheckAndNotifyStockMovementByIdAsync(int rmId, string transactionSource, string? refDocNo = null);

        /// <summary>
        /// Compiles and dispatches the daily consolidated stock threshold digest to all 'Stock' role users.
        /// </summary>
        Task SendDailyStockDigestAsync();
    }
}
