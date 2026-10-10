using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Stock
{
    public interface IRMDataAccess
    {
        Task<byte[]> GetMaterialPictureAsync(string materialId);
        Task<List<RMViewModel>> GetMaterialsAsync();
        Task<StockBreachSummaryDto> GetStockBreachSummaryAsync();
        Task<List<StockBreachItemDto>> GetStockBreachItemsAsync(string? breachType = null);
        Task<StockBreachItemDto?> GetMaterialStockStatusByCodeAsync(string materialId);
        Task<StockBreachItemDto?> GetMaterialStockStatusByIdAsync(int rmId);
    }
}
