
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Azure.Core;
using DataAccessLibrary;
using DataAccessLibrary.Models;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace CIP.Pages
{
    public partial class SaleBooking : ComponentBase
    {
        private Item item = new Item();
        private Customers customer = new Customers();
        public Item item_public { get { return item; } }
        
        private List<Customers> customers = new List<Customers>();

        private string successMessage = string.Empty;
        private decimal amount = 0;
        private string AutoNumber = string.Empty;


        private Customers SelectedCustomer;

        private async Task<IEnumerable<Customers>> GetCustomers(string searchText)
        {
            return await Task.FromResult(customers.Where(x => x.CustomerName.ToLower().Contains(searchText.ToLower())).ToList());
        }

        private async Task SelectedResultChanged(Customers customer)
        {
            SelectedCustomer = customer;
            Console.WriteLine($"Cust : {SelectedCustomer.CustCode} Name:{SelectedCustomer.CustomerName}");
        }
// NOW START WORKING FOR ITEMS
        private Material SelectedRM;
        private List<Material> materials = new List<Material>();

        private async Task<IEnumerable<Material>> GetRM(string searchText)
        {
            return await Task.FromResult(materials.Where(x => x.MaterialName.ToLower().Contains(searchText.ToLower())).ToList());
        }

        //private string Unit = string.Empty;
        //private int StockQty = 0;
        private async Task SelectedResultChanged_RM(Material material)
        {
            SelectedRM = material;
            Console.WriteLine($"RM : {SelectedRM.MaterialID} Name:{SelectedRM.MaterialName} Unit:{SelectedRM.Unit} StockQty:{SelectedRM.StockQty}");

            Unit = SelectedRM.Unit;
            StockQty = SelectedRM.StockQty;

        }

        private void Calculate()
        {
            amount = item.Quantity * item.Price;
        }

        private decimal CalculateAmount()
        {
            return item.Quantity * item.Price;
        }

        [Inject]
        private IItemData itemData { get; set; }
        
        [Inject]
        private ICustData custData { get; set; } // Inject service in Blazor

        private int maxBookingNumber;
        protected override async Task OnInitializedAsync()
        {
            await InitializeData();
        }
        private async Task SaveItem()
        {
            try
            {
                // Ensure RMDetails are set before saving
                if (Items == null || !Items.Any())
                {
                    throw new InvalidOperationException("No item added. Please add items before saving.");
                }

                // Set RMDetails before saving
                item.RMDetails = new List<RMDetail>();
                foreach (var rm in Items)
                {
                    item.RMDetails.Add(new RMDetail
                    {
                        ItemID = rm.ItemID,
                        Qty = rm.Qty,
                        GQty = rm.GQty
                        // Populate other properties as necessary
                    });
                }

                //await itemData.InsertItemAsync(item);
                item.CustCode = SelectedCustomer?.CustCode; // Assuming you save CustCode
                await itemData.InsertSaleBookingAsync(item);
                successMessage = "Item saved successfully!";
                //ResetForm();
            }
            catch (Exception ex)
            {
                successMessage = $"Error saving item: {ex.Message}";
            }
        }

        private List<Item> items = new List<Item>();
        private void ResetForm()
        {
            item = new Item();
            LogState("ResetForm");

            successMessage = string.Empty;
            amount = 0;
            SelectedCustomer = null;
            item.BookDescription = string.Empty;
            items.Clear();
            InitializeData().Wait(); // Wait for     the async call to complete
        }

        private async Task InitializeData()
        {
            // Initialize your item or any other data
            item.BookDT = DateTime.Today; // Sets the default date to today
            customers = await custData.GetCustomersAsync();
            // Manually adding sample customer data
            SelectedCustomer = new Customers("000", "Customer");

            materials = await itemData.GetMaterialsAsync();
            SelectedRM = new Material("000", "Material", "", 0);

            //AutoNumber = getNextSaleBookingInvoiceNo();
            await getNextSaleBookingInvoiceNo(item.BookDT);
            Console.WriteLine($"num : {item.BookDT}");
        }


        private void AddRMDetail()
        {
            item.RMDetails.Add(new RMDetail());
            LogState("AddRMDetail");
        }

        private void RemoveRMDetail(int index)
        {
            if (index >= 0 && index < item.RMDetails.Count)
            {
                item.RMDetails.RemoveAt(index);
                LogState("RemoveRMDetail");
            }
        }

        private void LogState(string method)
        {
            //Console.WriteLine($"{method} called. RMDetails Count: {item.RMDetails.Count}");
            Console.WriteLine($"Cust : {SelectedCustomer.CustCode} Name:{SelectedCustomer.CustomerName}");
            for (int i = 0; i < customer.CustDetails.Count; i++)
            {
                Console.WriteLine($"RMDetail[{i}] - RMQty: {item.RMDetails[i].RMCode}");
            }
        }

        [JSInvokable("UpdateRMCode")]
        public static Task UpdateRMCode(int index, string newValue, int x)
        {
            // Get the current instance of the component
            //var instance = new Test();
            //var instance = DotNetObjectReference.Create(this);
            // Call the non-static method
            //instance.UpdateRMCodeInstance(index, newValue);
            //item_public.RMDetails[index].RMCode = newValue;
            //InvokeAsync(() => StateHasChanged());
            //item.RMDetails[index].RMCode = newValue;
            Console.WriteLine($"Material ID : {newValue}");
            return Task.CompletedTask;
        }
        public void UpdateRMCodeInstance(int index, string newValue)
        {
            item.RMDetails[index].RMCode = newValue;
            Console.WriteLine($"Material ID : {newValue}");
            InvokeAsync(() => StateHasChanged());

        }
        private async void OnCountry(string countryId)
        {
            await InvokeAsync(StateHasChanged);
        }

        private DateTime todayDate = DateTime.Today;
        private string todayDateString => todayDate.ToString("yyyy-MM-dd");

        // Properties to bind to form inputs
        private string ItemName { get; set; }
        private string Unit { get; set; }
        private int StockQty { get; set; }
        private int Qty { get; set; }
        private int GQty { get; set; }

        // List to hold items temporarily
        private List<Item> Items { get; set; } = new List<Item>();
        private List<RMDetail> Items_rmd => item.RMDetails; // Access RMDetails for binding

        // Method to add an item to the table
        private int totalQuantity = 0;
        private void AddItem()
        {
            // CHECK ALREADY EXISTING ITEMS
            var newItem = new Item
            {
                //   ItemName = ItemName,               
                ItemName = SelectedRM.MaterialName,
                Unit = Unit,
                StockQty = StockQty,
                Qty = Qty,
                GQty = GQty,
                ItemID =SelectedRM.MaterialID,
            };

            // Check if database connection is established
            Console.WriteLine($"Material ID : {SelectedRM.MaterialID}");
            Items.Add(newItem);

            totalQuantity = totalQuantity + Qty;
            // Reset form fields
            ItemName = string.Empty;
            Unit = string.Empty;
            StockQty = 0;
            Qty = 0;
            GQty = 0;
            SelectedRM = null;

            // Refresh the UI
            StateHasChanged();
        }

        private void RemoveItem(Item item)
        {
            Items.Remove(item);
            StateHasChanged();
        }
        private int maxSBNumber;
        private DateTime selectedDate = DateTime.Today;
        private string BookInvoiceNo;

        private async Task getNextSaleBookingInvoiceNo(DateTime selectedDate)
        {
            maxSBNumber = await itemData.GetmaxBookingNumberAsync();

            string formattedBookingNumber = $"SBI-{selectedDate:yyMM}-{maxSBNumber + 1:000000}";
            item.BookInvoiceNo = formattedBookingNumber;
            Console.WriteLine($"num : {formattedBookingNumber}");
        }
        private async Task OnBookDTChanged(ChangeEventArgs e)
        {
            //selectedDate = newDate;
            await getNextSaleBookingInvoiceNo(item.BookDT);
            Console.WriteLine($"Date : {item.BookDT}");
        }

    }
}
