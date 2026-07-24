using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Stock
{
    public interface INewVendorDataAccess
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
