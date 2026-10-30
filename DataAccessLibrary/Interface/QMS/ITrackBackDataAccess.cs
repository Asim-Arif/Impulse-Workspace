using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.QMS;

namespace DataAccessLibrary.Interface.QMS
{
    public interface ITrackBackDataAccess
    {
        Task<List<string>> GetRecentCustomInvoicesAsync(string? searchTerm = null);
        Task<List<TrackBackInvoiceItemDto>> GetInvoiceItemsAsync(string invoiceNo);
        Task<List<TrackBackLotDto>> GetLotsByItemAsync(string invoiceNo, string compItemCode);
        Task<List<TrackBackMasterPODto>> GetMasterPOsByItemAsync(string invoiceNo, string compItemCode);
        Task<List<TrackBackOrderDto>> GetOrdersByItemAsync(string invoiceNo, string compItemCode);
        Task<List<string>> GetInvoiceMasterPONosAsync(string invoiceNo);
    }
}

