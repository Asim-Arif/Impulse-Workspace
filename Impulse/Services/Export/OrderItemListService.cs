using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Export;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public class OrderItemListService : IOrderItemListService
    {
        private readonly IOrderItemListDataAccess _dataAccess;

        public OrderItemListService(IOrderItemListDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<OrderItemListViewModel>> GetOrderItemsAsync(
            DateTime dtFrom, DateTime dtTo, string? custCode, string? country,
            int? groupID, string? orderNo, bool noDelivery, bool both, bool urgent)
            => _dataAccess.GetOrderItemsAsync(dtFrom, dtTo, custCode, country, groupID, orderNo, noDelivery, both, urgent);

        public Task<List<string>> GetCountriesForCustomerAsync(string custCode)
            => _dataAccess.GetCountriesForCustomerAsync(custCode);

        public Task<List<OrderNoLookupModel>> GetOrderNosForCustomerAsync(string custCode)
            => _dataAccess.GetOrderNosForCustomerAsync(custCode);

        public Task<List<OrderItemGroupLookupModel>> GetItemGroupsAsync()
            => _dataAccess.GetItemGroupsAsync();

        public Task<bool> UpdateDeliveryDateAsync(
            int id, DateTime newDeliveryDT, int prevDeliveryStatus, DateTime? prevDeliveryDT,
            string remarks, string userName, string machineName)
            => _dataAccess.UpdateDeliveryDateAsync(id, newDeliveryDT, prevDeliveryStatus, prevDeliveryDT, remarks, userName, machineName);
    }
}
