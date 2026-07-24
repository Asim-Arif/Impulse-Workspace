using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.AspNetCore.Components;
using System.Collections.Generic;

namespace Impulse.Components.Stock.NewRMTabs
{
    public partial class GeneralInfoTab : ComponentBase
    {
        [Parameter]
        public NewRMViewModel Model { get; set; } = default!;

        [Parameter]
        public List<string> Units { get; set; } = new();
    }
}
