using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.AspNetCore.Components;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Components.Stock.NewRMTabs
{
    public partial class UsedInTab : ComponentBase
    {
        [Parameter]
        public NewRMViewModel Model { get; set; } = default!;

        [Parameter]
        public List<ItemLookupModel> AvailableItems { get; set; } = new();

        public ItemLookupModel? selectedItem;

        private void OnItemSelected(ItemLookupModel? item)
        {
            selectedItem = item;
        }

        private Task<IEnumerable<ItemLookupModel>> SearchItems(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult(AvailableItems.Take(50));

            return Task.FromResult(AvailableItems
                .Where(x => 
                    (x.ItemID?.Contains(searchText, System.StringComparison.OrdinalIgnoreCase) ?? false) ||
                    (x.ItemName?.Contains(searchText, System.StringComparison.OrdinalIgnoreCase) ?? false))
                .Take(50));
        }

        private void AddItem()
        {
            if (selectedItem != null)
            {
                // Check if already added
                if (!Model.UsedInItems.Any(x => x.ItemID == selectedItem.ItemID && !x.IsDeleted))
                {
                    Model.UsedInItems.Add(new ItemRMCompViewModel
                    {
                        ItemID = selectedItem.ItemID,
                        ItemName = selectedItem.ItemName,
                        IsNew = true
                    });
                }
                
                selectedItem = null;
            }
        }

        private void RemoveItem(ItemRMCompViewModel item)
        {
            if (item.EntryID > 0)
            {
                // Existing item, mark as deleted for DB
                item.IsDeleted = true;
            }
            else
            {
                // New item, just remove from list
                Model.UsedInItems.Remove(item);
            }
        }
    }
}
