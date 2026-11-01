using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.FixedAssets;

namespace DataAccessLibrary.Interface.FixedAssets
{
    public interface IFixedAssetRegisterDataAccess
    {
        Task<FixedAssetRegisterFilterOptions> GetFilterOptionsAsync(CancellationToken cancellationToken = default);
        Task<(bool isSaved, List<FixedAssetRegisterRow> rows)> LoadRegisterDataAsync(DateTime dateFrom, DateTime dateTo, CancellationToken cancellationToken = default);
        Task<FixedAssetRegisterSaveResult> SaveRegisterAsync(DateTime dateFrom, DateTime dateTo, IReadOnlyList<FixedAssetRegisterRow> rows, string userName, string machineName, CancellationToken cancellationToken = default);
        Task<(bool success, string message)> DeleteRegisterAsync(DateTime dateFrom, CancellationToken cancellationToken = default);
        Task<int?> GetVoucherRefIdAsync(DateTime dateFrom, DateTime dateTo, CancellationToken cancellationToken = default);
    }
}

