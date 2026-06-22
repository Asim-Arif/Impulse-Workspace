
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Azure.Core;
using DataAccessLibrary;
using DataAccessLibrary.Models;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using DataAccessLibrary.Models;
using Microsoft.IdentityModel.Tokens;
using Microsoft.AspNetCore.Components.Forms;
using CIP.Controllers;

namespace CIP.Pages
{
        public partial class Banks : ComponentBase
        {
            private List<CIPBanks> banks = new();
            [Inject]
            private ICIPData ICIPBanks { get; set; }
            protected override async Task OnInitializedAsync()
            {
                banks = await ICIPBanks.getbanks();
            }

            private void NavigateToEntryForm(int bankId)
            {
                Navigation.NavigateTo($"/CIP_Bank/{bankId}");
            }

            private void ClosePage()
            {
                Navigation.NavigateTo("/"); // Redirect to the main or previous page
            }

        }
}
