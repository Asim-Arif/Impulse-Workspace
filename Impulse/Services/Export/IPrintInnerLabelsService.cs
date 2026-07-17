using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Company;

namespace Impulse.Services.Export
{
    public interface IPrintInnerLabelsService
    {
        Task<List<ItemLookupModel>> GetActiveItemsAsync();
        Task<List<string>> GetOrderNumbersForItemAsync(string itemId);
    }
}
