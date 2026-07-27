using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Stock
{
    public interface ISFTransactionDataAccess
    {
        Task<List<ItemViewModel>> GetArticlesAsync();
        Task<List<ProcessLookupModel>> GetProcessesAsync();
        Task<List<StoreLookupModel>> GetStoresAsync(string userName);
        Task<List<SFTransactionModel>> GetTransactionsAsync(SFTransactionFilterModel filter);
    }
}
