using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class FinishTransactionService : IFinishTransactionService
    {
        private readonly IFinishTransactionDataAccess _dataAccess;

        public FinishTransactionService(IFinishTransactionDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<FinishTransactionModel>> GetTransactionsAsync(FinishTransactionFilterModel filter) => _dataAccess.GetTransactionsAsync(filter);

        public Task<List<StoreLookupModel>> GetStoresAsync(string userName) => _dataAccess.GetStoresAsync(userName);

        public Task<List<ItemViewModel>> GetArticlesAsync() => _dataAccess.GetArticlesAsync();
    }
}
