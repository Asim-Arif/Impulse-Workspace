using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Production
{
    public class MakerRcvListService : IMakerRcvListService
    {
        private readonly IMakerRcvListDataAccess _dataAccess;

        public MakerRcvListService(IMakerRcvListDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<(List<MakerRcvListItem> Items, string ReportSql)> GetListAsync(MakerRcvListFilter filter) => _dataAccess.GetListAsync(filter);
        public Task<List<LookupItemInt>> GetMakersAsync() => _dataAccess.GetMakersAsync();
        public Task<List<LookupItemString>> GetCustomersAsync() => _dataAccess.GetCustomersAsync();
        public Task<List<LookupItemString>> GetItemCategoriesAsync() => _dataAccess.GetItemCategoriesAsync();
        public Task<List<LookupItemInt>> GetItemGroupsAsync() => _dataAccess.GetItemGroupsAsync();
        public Task<List<LookupItemString>> GetItemsAsync() => _dataAccess.GetItemsAsync();
        public Task<List<LookupItemInt>> GetProcessesAsync() => _dataAccess.GetProcessesAsync();
        public Task<bool> DeleteReceivingAsync(long vrdEntryId, string userName, string machineName) => _dataAccess.DeleteReceivingAsync(vrdEntryId, userName, machineName);
        public Task<bool> CloseLotAsync(string lotNo, string userName, string machineName) => _dataAccess.CloseLotAsync(lotNo, userName, machineName);
        public Task<bool> UpdateManualPTCNoAsync(string lotNo, string manualPTCNo) => _dataAccess.UpdateManualPTCNoAsync(lotNo, manualPTCNo);
        public Task<int> CheckIssuanceExistsAsync(long vrdEntryId) => _dataAccess.CheckIssuanceExistsAsync(vrdEntryId);
        public Task<bool> GetUserRightAsync(string rightName, string userName) => _dataAccess.GetUserRightAsync(rightName, userName);

        // Split / Transfer Lot lookups
        public Task<List<LookupItemString>> GetDistinctCustomerCodesAsync() => _dataAccess.GetDistinctCustomerCodesAsync();
        public Task<List<OrderLookupItem>> GetOrdersForCustomerAsync(string custCode, string? itemCode = null) => _dataAccess.GetOrdersForCustomerAsync(custCode, itemCode);
        public Task<List<LookupItemString>> GetArticlesForOrderAsync(string orderNo) => _dataAccess.GetArticlesForOrderAsync(orderNo);
        public Task<List<StoreLookupItem>> GetStoresAsync() => _dataAccess.GetStoresAsync();
        public Task<List<ShelfLookupItem>> GetShelvesByStoreAsync(int storeRefId) => _dataAccess.GetShelvesByStoreAsync(storeRefId);
        public Task<string> GetShelfRemarksAsync(string itemCode, int processId, int shelfRefId) => _dataAccess.GetShelfRemarksAsync(itemCode, processId, shelfRefId);

        // Split / Transfer Lot transactions
        public Task<bool> ChangeOrderNoAsync(ChangeOrderNoRequest request) => _dataAccess.ChangeOrderNoAsync(request);
        public Task<string> SplitLotAsync(SplitLotRequest request) => _dataAccess.SplitLotAsync(request);
        public Task<bool> TransferToSFStockAsync(TransferSFStockRequest request) => _dataAccess.TransferToSFStockAsync(request);

        // Inspection Data
        public Task<InspectionDataDto> GetInspectionDataAsync(long vrdEntryId) => _dataAccess.GetInspectionDataAsync(vrdEntryId);
        public Task<List<InspectionParameterItem>> GetProcessInspectionParametersAsync(int processId) => _dataAccess.GetProcessInspectionParametersAsync(processId);
        public Task<bool> SaveInspectionDataAsync(SaveInspectionRequest request) => _dataAccess.SaveInspectionDataAsync(request);
    }
}


