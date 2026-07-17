using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Company;

namespace Impulse.Services.Export
{
    public class PrintInnerLabelsService : IPrintInnerLabelsService
    {
        private readonly IPrintInnerLabelsDataAccess _dataAccess;

        public PrintInnerLabelsService(IPrintInnerLabelsDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<List<ItemLookupModel>> GetActiveItemsAsync()
        {
            return await _dataAccess.GetActiveItemsAsync();
        }

        public async Task<List<string>> GetOrderNumbersForItemAsync(string itemId)
        {
            return await _dataAccess.GetOrderNumbersForItemAsync(itemId);
        }
    }
}
