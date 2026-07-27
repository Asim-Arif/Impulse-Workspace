using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public interface IFinishTransactionService
    {
        Task<List<FinishTransactionModel>> GetTransactionsAsync(FinishTransactionFilterModel filter);
        Task<List<StoreLookupModel>> GetStoresAsync(string userName);
        Task<List<ItemViewModel>> GetArticlesAsync();
    }
}
