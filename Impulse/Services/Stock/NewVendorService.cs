using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class NewVendorService : INewVendorService
    {
        private readonly INewVendorDataAccess _newVendorDataAccess;

        public NewVendorService(INewVendorDataAccess newVendorDataAccess)
        {
            _newVendorDataAccess = newVendorDataAccess;
        }

        public Task<List<NewVendorViewModel>> GetVendorsAsync()
        {
            return _newVendorDataAccess.GetVendorsAsync();
        }

        public Task<NewVendorViewModel> GetVendorByIdAsync(int id)
        {
            return _newVendorDataAccess.GetVendorByIdAsync(id);
        }

        public Task<string> GetNextMakerNoAsync()
        {
            return _newVendorDataAccess.GetNextMakerNoAsync();
        }

        public Task<int> InsertVendorAsync(NewVendorViewModel model)
        {
            return _newVendorDataAccess.InsertVendorAsync(model);
        }

        public Task UpdateVendorAsync(NewVendorViewModel model)
        {
            return _newVendorDataAccess.UpdateVendorAsync(model);
        }

        public Task DeleteVendorAsync(int id)
        {
            return _newVendorDataAccess.DeleteVendorAsync(id);
        }

        public Task ToggleVendorStatusAsync(string accNo, bool makeActive)
        {
            return _newVendorDataAccess.ToggleVendorStatusAsync(accNo, makeActive);
        }
    }
}
