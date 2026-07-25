using DataAccessLibrary.Models.ViewModels.Stock;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public interface IStockLedgerService
    {
        Task<List<RMViewModel>> GetMaterialsAsync();
        Task<double> GetOpeningBalanceAsync(string materialId, DateTime fromDate);
        Task<List<StockLedgerTransactionModel>> GetTransactionsAsync(string materialId, DateTime fromDate, DateTime toDate);
    }
}
