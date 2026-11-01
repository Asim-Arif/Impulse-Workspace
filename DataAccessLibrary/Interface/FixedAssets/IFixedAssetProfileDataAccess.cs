using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.FixedAssets;

namespace DataAccessLibrary.Interface.FixedAssets
{
    public interface IFixedAssetProfileDataAccess
    {
        Task<FixedAssetLookupOptions> GetLookupOptionsAsync(CancellationToken cancellationToken = default);
        Task<FixedAssetProfileModel> NewAsync(CancellationToken cancellationToken = default);
        Task<IReadOnlyList<FixedAssetListRow>> GetListAsync(int? masterId, string? searchText, bool includeInactive, CancellationToken cancellationToken = default);
        Task<FixedAssetAdjustmentModel> NewAdjustmentAsync(string? assetId = null, CancellationToken cancellationToken = default);
        Task<int> SaveAdjustmentAsync(FixedAssetAdjustmentModel model, string userName, string machineName, CancellationToken cancellationToken = default);
        Task<FixedAssetMoveModel> NewMoveAsync(string? assetId = null, CancellationToken cancellationToken = default);
        Task SaveMoveAsync(FixedAssetMoveModel model, string userName, string machineName, CancellationToken cancellationToken = default);
        Task<FixedAssetDisposalModel> NewDisposalAsync(string? assetId = null, CancellationToken cancellationToken = default);
        Task SaveDisposalAsync(FixedAssetDisposalModel model, CancellationToken cancellationToken = default);
        Task<FixedAssetProfileModel?> GetAsync(int entryId, CancellationToken cancellationToken = default);
        Task<int> SaveAsync(FixedAssetProfileModel model, string userName, string machineName, CancellationToken cancellationToken = default);
        Task DeleteAssetAsync(int entryId, string assetId, CancellationToken cancellationToken = default);
        Task<int> SaveMasterCategoryAsync(string description, decimal percentRate, CancellationToken cancellationToken = default);
        Task SetActiveStatusAsync(int entryId, string assetId, bool isInactive, DateTime date, string remarks, string userName, string machineName, CancellationToken cancellationToken = default);
        Task<IReadOnlyList<FixedAssetRepairLogRow>> GetRepairLogsAsync(string assetId, CancellationToken cancellationToken = default);
    }
}

