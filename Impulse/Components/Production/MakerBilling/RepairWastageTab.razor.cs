using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.AspNetCore.Components;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Components.Production.MakerBilling
{
    public partial class RepairWastageTab : ComponentBase
    {
        [Parameter] public List<MakerRepairItem> Items { get; set; } = new List<MakerRepairItem>();
        [Parameter] public bool IsLoading { get; set; } = false;
        [Parameter] public EventCallback<decimal> OnRepairAmountChanged { get; set; }

        public string SearchLotText { get; set; } = string.Empty;

        public IEnumerable<MakerRepairItem> FilteredItems => string.IsNullOrWhiteSpace(SearchLotText)
            ? Items
            : Items.Where(x =>
                (x.LotNo?.Contains(SearchLotText.Trim(), StringComparison.OrdinalIgnoreCase) ?? false) ||
                (x.ItemCode?.Contains(SearchLotText.Trim(), StringComparison.OrdinalIgnoreCase) ?? false) ||
                (x.RecieptID?.Contains(SearchLotText.Trim(), StringComparison.OrdinalIgnoreCase) ?? false) ||
                (x.Process?.Contains(SearchLotText.Trim(), StringComparison.OrdinalIgnoreCase) ?? false) ||
                (x.WastageType?.Contains(SearchLotText.Trim(), StringComparison.OrdinalIgnoreCase) ?? false));

        public bool IsAllSelected => FilteredItems.Any() && FilteredItems.All(x => x.IsSelected);

        public decimal TotalSelectedRepairAmount => Items.Where(x => x.IsSelected).Sum(x => x.WastageAmt);

        public void ClearSearch()
        {
            SearchLotText = string.Empty;
        }

        public void SelectAll()
        {
            foreach (var item in FilteredItems) item.IsSelected = true;
            OnRepairAmountChanged.InvokeAsync(TotalSelectedRepairAmount);
        }

        public void UnselectAll()
        {
            foreach (var item in FilteredItems) item.IsSelected = false;
            OnRepairAmountChanged.InvokeAsync(TotalSelectedRepairAmount);
        }

        public void ToggleSelectAll(ChangeEventArgs e)
        {
            bool select = (bool)(e.Value ?? false);
            foreach (var item in FilteredItems) item.IsSelected = select;
            OnRepairAmountChanged.InvokeAsync(TotalSelectedRepairAmount);
        }

        public void OnSelectionChanged()
        {
            OnRepairAmountChanged.InvokeAsync(TotalSelectedRepairAmount);
        }
    }
}
