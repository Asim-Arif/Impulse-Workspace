using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.AspNetCore.Components;
using System.Collections.Generic;

namespace Impulse.Pages.Stock.VendorBilling.Tabs
{
    public partial class MaterialDeductionsTab : ComponentBase
    {
        [Parameter] public List<MaterialDeductionRowViewModel> Rows { get; set; } = new();
        [Parameter] public bool IsLoading { get; set; }
    }
}
