using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public interface INewVendorService
    {
        Task<List<NewVendorViewModel>> GetVendorsAsync();
        Task<NewVendorViewModel> GetVendorByIdAsync(int id);
        Task<string> GetNextMakerNoAsync();
        Task<int> InsertVendorAsync(NewVendorViewModel model);
        Task UpdateVendorAsync(NewVendorViewModel model);
        Task DeleteVendorAsync(int id);
        Task ToggleVendorStatusAsync(string accNo, bool makeActive);
    }
}
