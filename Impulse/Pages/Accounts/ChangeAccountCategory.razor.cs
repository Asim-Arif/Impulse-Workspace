using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Microsoft.AspNetCore.Components;

namespace Impulse.Pages.Accounts
{
    public partial class ChangeAccountCategory : ComponentBase
    {
        [Inject] private IChartOfAccountsDataAccess ChartOfAccountsDataAccess { get; set; }
        [Inject] private NavigationManager NavigationManager { get; set; }

        private bool IsLoading { get; set; } = true;
        private bool IsSaving { get; set; } = false;
        private string ErrorMessage { get; set; }
        private string SuccessMessage { get; set; }

        private List<ChartOfAccountsModel> AllCategories = new();
        private List<ChartOfAccountsModel> ChildAccounts = new();

        public ChartOfAccountsModel SelectedCategory { get; set; }
        public ChartOfAccountsModel SelectedAccount { get; set; }
        public ChartOfAccountsModel SelectedNewCategory { get; set; }

        private async Task OnCategoryChanged(ChartOfAccountsModel value)
        {
            SelectedCategory = value;
            SelectedAccount = null;
            SelectedNewCategory = null;
            NewAccountNo = string.Empty;
            await LoadChildAccountsAsync();
        }

        private void OnAccountChanged(ChartOfAccountsModel value)
        {
            SelectedAccount = value;
            SelectedNewCategory = null;
            NewAccountNo = string.Empty;
        }

        private async Task OnNewCategoryChanged(ChartOfAccountsModel value)
        {
            SelectedNewCategory = value;
            await CalculateNextAccountNoAsync();
        }

        public string NewAccountNo { get; set; }

        private bool CanSave => SelectedAccount != null && SelectedNewCategory != null && !string.IsNullOrEmpty(NewAccountNo) && !IsSaving;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                IsLoading = true;
                AllCategories = await ChartOfAccountsDataAccess.GetParentCategoriesAsync();
            }
            catch (Exception ex)
            {
                ErrorMessage = "Failed to load categories: " + ex.Message;
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task<IEnumerable<ChartOfAccountsModel>> SearchCategories(string searchText)
        {
            await Task.Delay(5); // Simulate async for Typeahead requirement
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return AllCategories;
            }
            
            return AllCategories.Where(c => 
                (c.AccNo?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false) ||
                (c.AccTitle?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false));
        }

        private async Task LoadChildAccountsAsync()
        {
            try
            {
                if (SelectedCategory == null)
                {
                    ChildAccounts.Clear();
                    return;
                }

                ChildAccounts = await ChartOfAccountsDataAccess.GetChildAccountsAsync(SelectedCategory.AccNo);
            }
            catch (Exception ex)
            {
                ErrorMessage = "Failed to load accounts: " + ex.Message;
            }
        }

        private async Task<IEnumerable<ChartOfAccountsModel>> SearchChildAccounts(string searchText)
        {
            await Task.Delay(5);
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return ChildAccounts;
            }

            return ChildAccounts.Where(c => 
                (c.AccNo?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false) ||
                (c.AccTitle?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false));
        }

        private async Task CalculateNextAccountNoAsync()
        {
            try
            {
                ErrorMessage = null;
                if (SelectedAccount == null || SelectedNewCategory == null)
                {
                    NewAccountNo = string.Empty;
                    return;
                }

                if (SelectedCategory?.AccNo == SelectedNewCategory.AccNo)
                {
                    ErrorMessage = "The old and new categories are the same.";
                    NewAccountNo = string.Empty;
                    return;
                }

                // Call the DAC to get the next account number. 
                // We use the existing account title, new category's type, and new category's ID.
                string newType = SelectedNewCategory.AccNo.Length >= 2 ? SelectedNewCategory.AccNo.Substring(0, 2) : SelectedNewCategory.AccNo;
                
                NewAccountNo = await ChartOfAccountsDataAccess.GetNextAccountNumberAsync(
                    acctype: newType, 
                    accountname: SelectedAccount.AccTitle, 
                    isparent: false, 
                    subaccof: SelectedNewCategory.AccNo);
            }
            catch (Exception ex)
            {
                ErrorMessage = "Failed to calculate the next Account Number: " + ex.Message;
            }
        }

        private async Task SaveChanges()
        {
            if (!CanSave) return;

            try
            {
                IsSaving = true;
                ErrorMessage = null;
                SuccessMessage = null;

                // In a real application, you might want to fetch actual user and machine names from the auth state/context
                string userName = "Admin"; 
                string machineName = "WebClient";

                await ChartOfAccountsDataAccess.ChangeAccountCategoryAsync(
                    oldAccNo: SelectedAccount.AccNo,
                    newAccNo: NewAccountNo,
                    newParentAccNo: SelectedNewCategory.AccNo,
                    userName: userName,
                    machineName: machineName);

                SuccessMessage = $"Successfully changed account {SelectedAccount.AccNo} to {NewAccountNo}.";
                
                // Reset form
                SelectedCategory = null;
            }
            catch (Exception ex)
            {
                ErrorMessage = "Failed to save changes: " + ex.Message;
            }
            finally
            {
                IsSaving = false;
            }
        }

        private void GoToIndexPage()
        {
            NavigationManager.NavigateTo("/financial");
        }
    }
}
