using DataAccessLibrary.Models.ViewModels.Stock;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Stock
{
    public interface IStockLedgerDataAccess
    {
        Task<List<RMViewModel>> GetMaterialsAsync();
        Task<double> GetOpeningBalanceAsync(string materialId, DateTime fromDate);
        Task<List<StockLedgerTransactionModel>> GetTransactionsAsync(string materialId, DateTime fromDate, DateTime toDate);
    }
}
