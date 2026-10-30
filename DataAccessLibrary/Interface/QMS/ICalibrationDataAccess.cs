using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.QMS;

namespace DataAccessLibrary.Interface.QMS
{
    public interface ICalibrationDataAccess
    {
        Task<List<CalibrationListItemModel>> GetCalibrationsAsync(CalibrationFilterDto filter);
        Task<CalibrationModel?> GetCalibrationByIdAsync(int entryId);
        Task<int> SaveCalibrationAsync(CalibrationModel model, string currentUserName, string? machineName);
        Task<bool> DeleteCalibrationAsync(int entryId);
        Task<List<FixedAssetLookupDto>> GetFixedAssetsAsync(bool requiresCalibrationOnly = true);
        Task<List<VendorLookupDto>> GetVendorsAsync();
    }
}

