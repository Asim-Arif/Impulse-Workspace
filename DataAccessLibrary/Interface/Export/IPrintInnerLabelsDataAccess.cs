using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Company;

namespace DataAccessLibrary.Interface.Export
{
    public interface IPrintInnerLabelsDataAccess
    {
        Task<List<ItemLookupModel>> GetActiveItemsAsync();
        Task<List<string>> GetOrderNumbersForItemAsync(string itemId);
    }
}
