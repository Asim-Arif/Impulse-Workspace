using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Stock;

namespace DataAccessLibrary.Interface.Stock
{
    public interface IVendGateRcvdListDataAccess
    {
        Task<List<object>> GetVendorsAsync();
        Task<List<object>> GetRMGroupsAsync();
        Task<List<VendGateRcvdListViewModel>> GetReceivingListAsync(DateTime fromDate, DateTime toDate, string vendorId, string poNo, string groupId, bool sampleOrder);
        Task DeleteReceivingAsync(string rcvId);
    }
}
