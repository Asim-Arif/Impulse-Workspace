using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Stock
{
    public class RMIssuanceService : IRMIssuanceService
    {
        private readonly IRMIssuanceDataAccess _dataAccess;

        public RMIssuanceService(IRMIssuanceDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<IEnumerable<MakerLookupModel>> GetMakersAsync()
        {
            return await _dataAccess.GetMakersAsync();
        }

        public async Task<IEnumerable<EmployeeLookupModel>> GetEmployeesAsync()
        {
            return await _dataAccess.GetEmployeesAsync();
        }

        public async Task<IEnumerable<RMLookupModel>> GetMaterialsAsync()
        {
            return await _dataAccess.GetMaterialsAsync();
        }

        public async Task<IEnumerable<RMLocationModel>> GetMaterialLocationsAsync(int rmid)
        {
            var locations = await _dataAccess.GetMaterialLocationsAsync(rmid);
            foreach (var loc in locations)
            {
                loc.Display = string.IsNullOrWhiteSpace(loc.BatchNo) 
                    ? $"{loc.Store} ({loc.ShelfQty})" 
                    : $"Batch:{loc.BatchNo} Lot:{loc.LotNo} {loc.Store} ({loc.ShelfQty})";
            }
            return locations;
        }

        public async Task<RMDetailsModel> GetRMDetailsAsync(int rmid)
        {
            return await _dataAccess.GetRMDetailsAsync(rmid);
        }

        public async Task<IEnumerable<RMHistoryRowModel>> GetIssuanceHistoryAsync(bool isMaker, string recipientId, DateTime fromDate, DateTime toDate)
        {
            return await _dataAccess.GetIssuanceHistoryAsync(isMaker, recipientId, fromDate, toDate);
        }

        public async Task<IEnumerable<IssuanceReceiptLookupModel>> GetProductionIssuancesAsync(int vendorId)
        {
            return await _dataAccess.GetProductionIssuancesAsync(vendorId);
        }

        public async Task<IEnumerable<DataAccessLibrary.Models.ViewModels.Stock.RMLookupModel>> GetMaterialsByIssuanceAsync(int entryId)
        {
            return await _dataAccess.GetMaterialsByIssuanceAsync(entryId);
        }

        public async Task<string> SaveIssuanceTransactionAsync(RMIssuanceSaveRequest request)
        {
            return await _dataAccess.SaveIssuanceTransactionAsync(request);
        }
    }
}
