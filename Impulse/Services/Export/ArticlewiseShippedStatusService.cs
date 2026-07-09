using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public class ArticlewiseShippedStatusService : IArticlewiseShippedStatusService
    {
        private readonly IArticlewiseShippedStatusDataAccess _dataAccess;

        public ArticlewiseShippedStatusService(IArticlewiseShippedStatusDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<IEnumerable<ArticleShippedInvoiceModel>> GetShippedInvoicesAsync(string itemCode, string custCode)
        {
            return _dataAccess.GetShippedInvoicesAsync(itemCode, custCode);
        }

        public Task<IEnumerable<ArticleUnshippedOrderModel>> GetUnshippedOrdersAsync(string itemCode, string custCode)
        {
            return _dataAccess.GetUnshippedOrdersAsync(itemCode, custCode);
        }

        public Task<IEnumerable<ArticleCustomerPriceModel>> GetCustomerPricesAsync(string itemCode)
        {
            return _dataAccess.GetCustomerPricesAsync(itemCode);
        }

        public Task<IEnumerable<ArticleQuotationModel>> GetQuotationsAsync(string itemCode)
        {
            return _dataAccess.GetQuotationsAsync(itemCode);
        }

        public Task UpdateCustomerPriceAsync(long entryId, string columnName, double newValue, string userName)
        {
            return _dataAccess.UpdateCustomerPriceAsync(entryId, columnName, newValue, userName);
        }
    }
}
