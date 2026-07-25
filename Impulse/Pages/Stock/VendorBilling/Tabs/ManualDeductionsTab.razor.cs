using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Pages.Stock.VendorBilling.Tabs
{
    public partial class ManualDeductionsTab : ComponentBase
    {
        [Parameter] public List<ManualDeductionRowViewModel> Rows { get; set; } = new();
        [Parameter] public EventCallback OnDataChanged { get; set; }

        private string NewDescription { get; set; } = string.Empty;
        private decimal NewAmount { get; set; } = 0;

        private async Task AddRow()
        {
            if (string.IsNullOrWhiteSpace(NewDescription)) return;
            Rows.Add(new ManualDeductionRowViewModel
            {
                Description = NewDescription.Trim(),
                Amount = NewAmount
            });
            NewDescription = string.Empty;
            NewAmount = 0;
            await OnDataChanged.InvokeAsync();
        }

        private async Task RemoveRow(ManualDeductionRowViewModel row)
        {
            Rows.Remove(row);
            await OnDataChanged.InvokeAsync();
        }

        private async Task OnInputKeyDown(KeyboardEventArgs e)
        {
            if (e.Key == "Enter") await AddRow();
        }
    }
}
