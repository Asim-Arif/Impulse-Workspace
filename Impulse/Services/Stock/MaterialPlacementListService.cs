using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class MaterialPlacementListService : IMaterialPlacementListService
    {
        private readonly IMaterialPlacementListDataAccess _dataAccess;

        public MaterialPlacementListService(IMaterialPlacementListDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<IEnumerable<PlacedMaterialItemViewModel>> GetPlacedMaterialsListAsync(DateTime fromDate, DateTime toDate, string vendorId, string materialId)
        {
            return _dataAccess.GetPlacedMaterialsListAsync(fromDate, toDate, vendorId, materialId);
        }

        public Task<IEnumerable<object>> GetVendorsAsync()
        {
            return _dataAccess.GetVendorsAsync();
        }

        public Task<IEnumerable<object>> GetMaterialsAsync()
        {
            return _dataAccess.GetMaterialsAsync();
        }
    }
}
