using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.AspNetCore.Components;
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

        public bool IsAllSelected { get; set; } = true;

        public decimal TotalSelectedRepairAmount => Items.Where(x => x.IsSelected).Sum(x => x.WastageAmt);

        public void SelectAll()
        {
            foreach (var item in Items) item.IsSelected = true;
            IsAllSelected = true;
            OnRepairAmountChanged.InvokeAsync(TotalSelectedRepairAmount);
        }

        public void UnselectAll()
        {
            foreach (var item in Items) item.IsSelected = false;
            IsAllSelected = false;
            OnRepairAmountChanged.InvokeAsync(TotalSelectedRepairAmount);
        }

        public void ToggleSelectAll(ChangeEventArgs e)
        {
            IsAllSelected = (bool)(e.Value ?? false);
            foreach (var item in Items) item.IsSelected = IsAllSelected;
            OnRepairAmountChanged.InvokeAsync(TotalSelectedRepairAmount);
        }

        public void OnSelectionChanged()
        {
            IsAllSelected = Items.Any() && Items.All(x => x.IsSelected);
            OnRepairAmountChanged.InvokeAsync(TotalSelectedRepairAmount);
        }
    }
}
