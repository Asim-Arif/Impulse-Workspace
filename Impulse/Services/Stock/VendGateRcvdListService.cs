using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Stock;
using DataAccessLibrary.Models.ViewModels.Stock;

namespace Impulse.Services.Stock
{
    public class VendGateRcvdListService : IVendGateRcvdListService
    {
        private readonly IVendGateRcvdListDataAccess _dataAccess;

        public VendGateRcvdListService(IVendGateRcvdListDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<List<object>> GetVendorsAsync()
        {
            return await _dataAccess.GetVendorsAsync();
        }

        public async Task<List<object>> GetRMGroupsAsync()
        {
            return await _dataAccess.GetRMGroupsAsync();
        }

        public async Task<List<VendGateRcvdListViewModel>> GetReceivingListAsync(DateTime fromDate, DateTime toDate, string vendorId, string poNo, string groupId, bool sampleOrder)
        {
            return await _dataAccess.GetReceivingListAsync(fromDate, toDate, vendorId, poNo, groupId, sampleOrder);
        }

        public async Task DeleteReceivingAsync(string rcvId)
        {
            await _dataAccess.DeleteReceivingAsync(rcvId);
        }
    }
}
