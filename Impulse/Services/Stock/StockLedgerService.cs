using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class StockLedgerService : IStockLedgerService
    {
        private readonly IStockLedgerDataAccess _dataAccess;

        public StockLedgerService(IStockLedgerDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<RMViewModel>> GetMaterialsAsync()
        {
            return _dataAccess.GetMaterialsAsync();
        }

        public Task<double> GetOpeningBalanceAsync(string materialId, DateTime fromDate)
        {
            return _dataAccess.GetOpeningBalanceAsync(materialId, fromDate);
        }

        public Task<List<StockLedgerTransactionModel>> GetTransactionsAsync(string materialId, DateTime fromDate, DateTime toDate)
        {
            return _dataAccess.GetTransactionsAsync(materialId, fromDate, toDate);
        }
    }
}
