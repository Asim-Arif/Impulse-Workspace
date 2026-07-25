using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Company;

namespace DataAccessLibrary.Interface.Stock
{
    public interface IVendOrderDataAccess
    {
        Task<VendOrderViewModel> GetVendOrderAsync(string orderNo);
        Task<string> InsertVendOrderAsync(VendOrderViewModel model);
        Task<bool> UpdateVendOrderAsync(VendOrderViewModel model);
        Task<bool> DeleteVendOrderAsync(string orderNo);
        
        Task<string> GetNextOrderNoAsync();
        Task<IEnumerable<RMPOListViewModel>> GetRMPOListAsync(RMPOListSearchFilter filter);
        Task UpdatePDFAttachmentAsync(string orderNo, byte[] pdfData);
    }
}
