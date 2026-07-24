using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels.Stock;
using Impulse.Services;
using Impulse.Services.Stock;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.JSInterop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Stock
{
    public partial class NewVendor : ComponentBase
    {
        [Inject]
        public INewVendorService NewVendorService { get; set; } = default!;

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

        [SupplyParameterFromQuery]
        public string? ReturnUrl { get; set; }

        public NewVendorViewModel Model { get; set; } = new();

        public List<ChartOfAccountsModel> Categories { get; set; } = new();

        private ChartOfAccountsModel _selectedCategory;
        public ChartOfAccountsModel SelectedCategory
        {
            get => _selectedCategory;
            set
            {
                _selectedCategory = value;
                Model.ParentAccNo = value?.AccNo ?? string.Empty;
            }
        }

        private bool isLoading = true;
        private bool isSaving = false;

        protected override async Task OnInitializedAsync()
        {
            isLoading = true;
            try
            {
                Categories = await ChartOfAccountsDataAccess.GetParentCategoriesAsync();

                if (Id != 0)
                {
                    Model = await NewVendorService.GetVendorByIdAsync(Id);
                    
                    if (!string.IsNullOrEmpty(Model.ParentAccNo))
                    {
                        _selectedCategory = Categories.FirstOrDefault(c => c.AccNo == Model.ParentAccNo);
                    }
                }
                else
                {
                    Model.MakerNo = await NewVendorService.GetNextMakerNoAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", $"Error loading data: {ex.Message}");
            }
            finally
            {
                isLoading = false;
            }
        }

        private async Task<IEnumerable<ChartOfAccountsModel>> SearchCategories(string searchText)
        {
            await Task.Delay(5); // Simulate async work
            if (string.IsNullOrWhiteSpace(searchText))
                return Categories;

            return Categories.Where(c => c.AccTitle.Contains(searchText, StringComparison.OrdinalIgnoreCase) || 
                                         c.AccNo.Contains(searchText, StringComparison.OrdinalIgnoreCase));
        }

        private async Task SaveVendor()
        {
            if (string.IsNullOrWhiteSpace(Model.VendorName))
            {
                NotificationService.ShowWarning("Validation", "Please enter a Vendor Name.");
                return;
            }

            if (string.IsNullOrWhiteSpace(Model.ParentAccNo))
            {
                NotificationService.ShowWarning("Validation", "Please select a Category.");
                return;
            }

            isSaving = true;
            try
            {
                if (Id == 0)
                {
                    int newId = await NewVendorService.InsertVendorAsync(Model);
                    NotificationService.ShowSuccess("Saved", "Vendor and Account created successfully.");
                }
                else
                {
                    await NewVendorService.UpdateVendorAsync(Model);
                    NotificationService.ShowSuccess("Updated", "Vendor updated successfully.");
                }

                NavigationManager.NavigateTo(string.IsNullOrEmpty(ReturnUrl) ? "/stock" : ReturnUrl);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", $"Error saving vendor: {ex.Message}");
            }
            finally
            {
                isSaving = false;
            }
        }

        private void GoBack()
        {
            NavigationManager.NavigateTo(string.IsNullOrEmpty(ReturnUrl) ? "/stock" : ReturnUrl);
        }

        private async Task LoadVendorPic(InputFileChangeEventArgs e)
        {
            var file = e.File;
            if (file != null)
            {
                var format = "image/png";
                var resizedImageFile = await file.RequestImageFileAsync(format, 400, 400);
                var buffer = new byte[resizedImageFile.Size];
                await resizedImageFile.OpenReadStream().ReadAsync(buffer);
                Model.VendorPic = $"data:{format};base64,{Convert.ToBase64String(buffer)}";
            }
        }

        private async Task LoadVendorSig(InputFileChangeEventArgs e)
        {
            var file = e.File;
            if (file != null)
            {
                var format = "image/png";
                var resizedImageFile = await file.RequestImageFileAsync(format, 400, 400);
                var buffer = new byte[resizedImageFile.Size];
                await resizedImageFile.OpenReadStream().ReadAsync(buffer);
                Model.VendorSig = $"data:{format};base64,{Convert.ToBase64String(buffer)}";
            }
        }
    }
}
