using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using Impulse.Services;
using Impulse.Services.Stock;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Stock
{
    public partial class NewRM : ComponentBase
    {
        [Inject]
        public INewRMService NewRMService { get; set; } = default!;

        [Inject]
        public IRMGroupsService RMGroupsService { get; set; } = default!;

        [Inject]
        public IChartOfAccountsDataAccess ChartOfAccountsDataAccess { get; set; } = default!;

        [Inject]
        public NavigationManager NavigationManager { get; set; } = default!;

        [Inject]
        public IJSRuntime JSRuntime { get; set; } = default!;

        [Inject]
        public INotificationService NotificationService { get; set; } = default!;

        [Parameter]
        public int Id { get; set; }

        [Parameter]
        public int CopyFromId { get; set; }

        public NewRMViewModel Model { get; set; } = new();

        public List<RMGroupViewModel> Categories { get; set; } = new();
        public List<ChartOfAccountsModel> SaleAccounts { get; set; } = new();
        public List<ChartOfAccountsModel> SaleRAccounts { get; set; } = new();
        public List<ChartOfAccountsModel> AssetAccounts { get; set; } = new();
        public List<ItemLookupModel> AvailableItems { get; set; } = new();
        public List<string> Units { get; set; } = new();

        private string activeTab = "general";
        private bool isLoading = true;
        private bool isSaving = false;

        protected override async Task OnInitializedAsync()
        {
            isLoading = true;
            try
            {
                await LoadInitialData();

                if (Id > 0)
                {
                    Model = await NewRMService.GetRMAsync(Id);
                    if (Model == null)
                    {
                        Model = new NewRMViewModel();
                    }
                }
                else if (CopyFromId > 0)
                {
                    // Legacy frmNewRM.AddNew(p_bCopy:=True): load the source record, then reset
                    // the identity so saving inserts a new material (Device Code stays editable).
                    Model = await NewRMService.GetRMAsync(CopyFromId);
                    if (Model == null)
                    {
                        Model = new NewRMViewModel();
                    }
                    else
                    {
                        Model.RMID = 0;
                        Model.TechnicalDrawing = null; // legacy does not copy the technical drawing
                        foreach (var u in Model.UsedInItems)
                        {
                            u.EntryID = 0; // legacy resets EntryID so they are inserted for the new RM
                            u.IsNew = true;
                        }
                    }
                }
                else
                {
                    Model = new NewRMViewModel();
                }
            }
            catch (Exception ex)
            {
                // Handle exception / toast
            }
            finally
            {
                isLoading = false;
            }
        }

        private async Task LoadInitialData()
        {
            var groups = await RMGroupsService.GetRMGroupsAsync();
            Categories = groups.ToList();

            // Load Accounts for dropdowns
            var allAccounts = await ChartOfAccountsDataAccess.GetChartOfAccountsList(false);
            SaleAccounts = allAccounts.Where(x => x.AccNo != null && (x.AccNo.StartsWith("4") || x.AccNo.StartsWith("3"))).ToList(); // Common for sales/revenue
            SaleRAccounts = allAccounts.Where(x => x.AccNo != null && (x.AccNo.StartsWith("4") || x.AccNo.StartsWith("3") || x.AccNo.StartsWith("5"))).ToList();
            AssetAccounts = allAccounts.Where(x => x.AccNo != null && x.AccNo.StartsWith("1")).ToList();

            // Load Items for Used In
            var items = await NewRMService.GetAvailableItemsAsync();
            AvailableItems = items;

            // Setup basic units list based on old code typically
            Units = new List<string> { "KGs", "Ltr", "Nos", "Mtr", "Ft", "Pair", "Set", "Pcs", "Gm" };
        }

        private void SetActiveTab(string tabName)
        {
            activeTab = tabName;
        }

        private async Task SaveData()
        {
            isSaving = true;
            try
            {
                // Validate
                if (Model.GroupID == 0)
                {
                    NotificationService.ShowWarning("Validation", "Please select a Category (Group).");
                    return;
                }
                if (string.IsNullOrWhiteSpace(Model.RMID1))
                {
                    NotificationService.ShowWarning("Validation", "Please enter a Device Code (RMID1).");
                    return;
                }

                await NewRMService.SaveRMAsync(Model);

                NotificationService.ShowSuccess("Saved", "Raw material saved successfully.");

                NavigationManager.NavigateTo("/stock");
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", $"Error saving: {ex.Message}");
            }
            finally
            {
                isSaving = false;
            }
        }

        private void GoBack()
        {
            NavigationManager.NavigateTo("/stock");
        }
    }
}
