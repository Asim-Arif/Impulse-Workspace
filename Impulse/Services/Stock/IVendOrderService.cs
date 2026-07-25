using DataAccessLibrary.Models.ViewModels.Stock;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public interface IVendOrderService
    {
        Task<VendOrderViewModel> GetVendOrderAsync(string orderNo);
        Task<string> SaveVendOrderAsync(VendOrderViewModel model);
        Task<bool> DeleteVendOrderAsync(string orderNo);
        Task<string> GetNextOrderNoAsync();
        Task<IEnumerable<RMPOListViewModel>> GetRMPOListAsync(RMPOListSearchFilter filter);
        Task UpdatePDFAttachmentAsync(string orderNo, byte[] pdfData);
    }
}
