using DataAccessLibrary.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface
{
    internal interface ICustData
    {

    }
}

namespace DataAccessLibrary
{
    public interface ICustData
    {
        Task<List<Customers>> GetCustomersAsync(); // (Optional) Retrieve all Customers

        //Task InsertSalesBookingAsync(Item item);            // Method to save the item
        //Task<Item> GetItemByIdAsync(string itemid);        // (Optional) Retrieve item by ID
        //Task<IEnumerable<Item>> GetItemsAsync();    // (Optional) Retrieve all items
    }
}