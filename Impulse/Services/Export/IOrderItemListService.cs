using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services.Export
{
    public interface IOrderItemListService
    {
        Task<List<OrderItemListViewModel>> GetOrderItemsAsync(
            DateTime dtFrom,
            DateTime dtTo,
            string? custCode,
            string? country,
            int? groupID,
            string? orderNo,
            bool noDelivery,
            bool both,
            bool urgent);

        Task<List<string>> GetCountriesForCustomerAsync(string custCode);
        Task<List<OrderNoLookupModel>> GetOrderNosForCustomerAsync(string custCode);
        Task<List<OrderItemGroupLookupModel>> GetItemGroupsAsync();

        Task<bool> UpdateDeliveryDateAsync(
            int id,
            DateTime newDeliveryDT,
            int prevDeliveryStatus,
            DateTime? prevDeliveryDT,
            string remarks,
            string userName,
            string machineName);
    }
}
