using DataAccessLibrary.Models.ViewModels.Stock;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Stock
{
    public interface IMaterialPlacementListDataAccess
    {
        Task<IEnumerable<PlacedMaterialItemViewModel>> GetPlacedMaterialsListAsync(DateTime fromDate, DateTime toDate, string vendorId, string materialId);
        Task<IEnumerable<object>> GetVendorsAsync();
        Task<IEnumerable<object>> GetMaterialsAsync();
    }
}
