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
using Impulse.Controllers;

namespace Impulse.Pages.Accounts
{
        public partial class FinancialStatementReports : ComponentBase
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

        //private void ClosePage() => Navigation.NavigateTo("/");
        //private void NavigateToEntryForm(int id) { /* Your Navigation Logic */ }

        private List<ReportDetail> reportList = new()
        {
            new(1, "Balance Sheet", "fas fa-file-invoice-dollar", "bg-danger bg-opacity-10", "text-danger", "View Assets, Liabilities & Equity"),
            new(2, "Profit and Loss", "fas fa-chart-line", "bg-primary bg-opacity-10", "text-primary", "Revenue and Expense Summary"),
            new(3, "Financial Notes", "fas fa-file-contract", "bg-success bg-opacity-10", "text-success", "Statement Detailed Breakdown"),
            new(4, "Sub-Heads Report", "fas fa-sitemap", "bg-warning bg-opacity-10", "text-warning", "Print Sub Heads Report"),
            new(5, "Depreciation", "fas fa-history", "bg-secondary bg-opacity-10", "text-secondary", "Fixed Asset Tracking"),
            new(6, "New Accounts", "fas fa-calendar-check", "bg-dark bg-opacity-10", "text-dark", "Datewise New Added Heads"),
            new(7, "Print Attachments", "fas fa-paperclip", "bg-info bg-opacity-10", "text-info", "View attached supporting documents"),
            new(8, "Capital WIP", "fas fa-tools", "bg-warning bg-opacity-10", "text-warning", "Print Capital Work In Progress report"),
            new(9, "Cash Flow Statement", "fas fa-exchange-alt", "bg-success bg-opacity-10", "text-success", "Print Cash Flow Statement")
        };

        public record ReportDetail(int Id, string Title, string Icon, string BgClass, string TextClass, string Description);

    }
}
