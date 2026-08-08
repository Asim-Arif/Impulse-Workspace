using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.AspNetCore.Components;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Components.Production.MakerBilling
{
    public partial class ManualDeductionsTab : ComponentBase
    {
        [Parameter] public List<MakerManualDedItem> Items { get; set; } = new List<MakerManualDedItem>();
        [Parameter] public EventCallback<decimal> OnManualAmountChanged { get; set; }

        public string NewDescription { get; set; } = string.Empty;
        public decimal NewAmount { get; set; } = 0;

        public decimal TotalManualAmount => Items.Sum(x => x.Amount);

        public void AddManualDeduction()
        {
            if (string.IsNullOrWhiteSpace(NewDescription) || NewAmount <= 0) return;

            Items.Add(new MakerManualDedItem
            {
                Description = NewDescription.Trim(),
                Amount = NewAmount
            });

            NewDescription = string.Empty;
            NewAmount = 0;

            OnManualAmountChanged.InvokeAsync(TotalManualAmount);
        }

        public void RemoveItem(MakerManualDedItem item)
        {
            Items.Remove(item);
            OnManualAmountChanged.InvokeAsync(TotalManualAmount);
        }
    }
}
