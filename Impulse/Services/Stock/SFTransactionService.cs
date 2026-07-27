using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class SFTransactionService : ISFTransactionService
    {
        private readonly ISFTransactionDataAccess _dataAccess;

        public SFTransactionService(ISFTransactionDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<ItemViewModel>> GetArticlesAsync() => _dataAccess.GetArticlesAsync();

        public Task<List<ProcessLookupModel>> GetProcessesAsync() => _dataAccess.GetProcessesAsync();

        public Task<List<StoreLookupModel>> GetStoresAsync(string userName) => _dataAccess.GetStoresAsync(userName);

        public Task<List<SFTransactionModel>> GetTransactionsAsync(SFTransactionFilterModel filter) => _dataAccess.GetTransactionsAsync(filter);
    }
}
