using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public interface IArticlewiseShippedStatusService
    {
        Task<IEnumerable<ArticleShippedInvoiceModel>> GetShippedInvoicesAsync(string itemCode, string custCode);
        Task<IEnumerable<ArticleUnshippedOrderModel>> GetUnshippedOrdersAsync(string itemCode, string custCode);
        Task<IEnumerable<ArticleCustomerPriceModel>> GetCustomerPricesAsync(string itemCode);
        Task<IEnumerable<ArticleQuotationModel>> GetQuotationsAsync(string itemCode);
        
        Task UpdateCustomerPriceAsync(long entryId, string columnName, double newValue, string userName);
    }
}
