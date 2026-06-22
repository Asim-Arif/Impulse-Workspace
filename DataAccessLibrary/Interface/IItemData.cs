using DataAccessLibrary.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface
{
    internal interface IItemData
    {

    }
}

namespace DataAccessLibrary
{
    public interface IItemData
    {
        Task<List<Material>> GetMaterialsAsync();
        Task InsertItemAsync(Item item);            // Method to save the item
        Task<Item> GetItemByIdAsync(string itemid);        // (Optional) Retrieve item by ID
        Task<IEnumerable<Item>> GetItemsAsync();    // (Optional) Retrieve all items
        Task InsertSaleBookingAsync(Item item);            // Method to save the item
        Task<int> GetmaxBookingNumberAsync();
        //Task<List<SBList>> GetSalesBookingListAsync();
        Task<List<Item>> GetSalesBookingListAsync();
    }
}