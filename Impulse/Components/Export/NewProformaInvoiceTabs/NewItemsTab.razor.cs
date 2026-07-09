using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Services.Export;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;

namespace Impulse.Components.Export.NewProformaInvoiceTabs
{
    public partial class NewItemsTab
    {
        [Inject] private IProformaInvoiceService ProformaService { get; set; } = default!;

        [Parameter] public int CompanyRefID { get; set; }
        [Parameter] public string CustCode { get; set; } = string.Empty;
        [Parameter] public string Country { get; set; } = string.Empty;
        
        [Parameter] public EventCallback<List<OrderItemForProformaModel>> OnAddItems { get; set; }

        public List<string> AvailableOrders { get; set; } = new();
        public string SelectedOrderNo { get; set; } = string.Empty;
        
        private List<OrderItemForProformaModel> OrderItems { get; set; } = new();
        private List<OrderItemForProformaModel> FilteredOrderItems { get; set; } = new();

        private string SearchTerm { get; set; } = string.Empty;
        private bool SelectAllChecked { get; set; }

        private bool HasItems => FilteredOrderItems != null && FilteredOrderItems.Any();
        private bool HasSelectedItems => FilteredOrderItems != null && FilteredOrderItems.Any(i => i.IsSelected);

        private int EditingQtyIdx = -1;

        private int _prevCompanyId;
        private string _prevCustCode = string.Empty;
        private string _prevCountry = string.Empty;

        protected override async Task OnParametersSetAsync()
        {
            if (CompanyRefID != _prevCompanyId || CustCode != _prevCustCode || Country != _prevCountry)
            {
                _prevCompanyId = CompanyRefID;
                _prevCustCode = CustCode;
                _prevCountry = Country;
                await LoadAvailableOrders(CompanyRefID, CustCode, Country);
            }
        }

        public async Task LoadAvailableOrders(int companyId, string custCode, string country)
        {
            if (companyId > 0 && !string.IsNullOrEmpty(custCode) && !string.IsNullOrEmpty(country))
            {
                AvailableOrders = await ProformaService.GetOrdersForProformaAsync(companyId, custCode, country);
                SelectedOrderNo = string.Empty;
                OrderItems.Clear();
                FilterItems();
            }
            else
            {
                AvailableOrders.Clear();
                SelectedOrderNo = string.Empty;
                OrderItems.Clear();
                FilterItems();
            }
        }

        private async Task OnOrderNoChanged()
        {
            if (!string.IsNullOrEmpty(SelectedOrderNo))
            {
                OrderItems = await ProformaService.GetOrderItemsForProformaAsync(SelectedOrderNo);
                FilterItems();
                SelectAllChecked = false;
            }
            else
            {
                OrderItems.Clear();
                FilterItems();
                SelectAllChecked = false;
            }
        }

        private void FilterItems()
        {
            if (string.IsNullOrWhiteSpace(SearchTerm))
            {
                FilteredOrderItems = OrderItems.ToList();
            }
            else
            {
                string term = SearchTerm.ToLower();
                FilteredOrderItems = OrderItems.Where(i => 
                    i.ItemCode.ToLower().Contains(term) || 
                    i.Description.ToLower().Contains(term)).ToList();
            }
        }

        private void OnSearchKeyUp(KeyboardEventArgs e)
        {
            FilterItems();
        }

        private void OnSelectAllChanged()
        {
            if (FilteredOrderItems != null)
            {
                foreach (var item in FilteredOrderItems)
                {
                    item.IsSelected = SelectAllChecked;
                }
            }
        }

        private void OnItemSelectionChanged()
        {
            if (FilteredOrderItems != null && FilteredOrderItems.Any())
            {
                SelectAllChecked = FilteredOrderItems.All(i => i.IsSelected);
            }
            else
            {
                SelectAllChecked = false;
            }
        }

        private void StartEditQty(int idx)
        {
            EditingQtyIdx = idx;
        }

        private void CancelEdit()
        {
            EditingQtyIdx = -1;
        }

        private void CommitQty(int idx, string? val)
        {
            if (idx >= 0 && idx < FilteredOrderItems.Count)
            {
                if (int.TryParse(val, out int qty) && qty > 0)
                {
                    if (qty > FilteredOrderItems[idx].InvQty)
                    {
                        FilteredOrderItems[idx].QtyToAdd = FilteredOrderItems[idx].InvQty;
                    }
                    else
                    {
                        FilteredOrderItems[idx].QtyToAdd = qty;
                    }
                }
                else
                {
                    FilteredOrderItems[idx].QtyToAdd = 1;
                }
            }
            CancelEdit();
        }

        private async Task AddSelectedItems()
        {
            var selectedItems = FilteredOrderItems.Where(i => i.IsSelected).ToList();
            if (selectedItems.Any())
            {
                await OnAddItems.InvokeAsync(selectedItems);
                
                foreach (var item in selectedItems)
                {
                    item.IsSelected = false;
                }
                SelectAllChecked = false;
            }
        }

        private async Task AddAllItems()
        {
            if (FilteredOrderItems.Any())
            {
                await OnAddItems.InvokeAsync(FilteredOrderItems.ToList());
                
                foreach (var item in FilteredOrderItems)
                {
                    item.IsSelected = false;
                }
                SelectAllChecked = false;
            }
        }
    }
}
