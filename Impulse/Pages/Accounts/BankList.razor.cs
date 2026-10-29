using BlazorContextMenu;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Accounts
{
    public partial class BankList : ComponentBase, IDisposable
    {
        [SupplyParameterFromQuery]
        public string? ReturnUrl { get; set; }

        [Inject]
        private IBankListDataAccess BankListDataAccess { get; set; } = default!;

        [Inject]
        private IAuditService _auditService { get; set; } = default!;

        [Inject]
        private INotificationService NotificationServiceManager { get; set; } = default!;

        [Inject]
        private DialogService DialogService { get; set; } = default!;

        [Inject]
        private NavigationManager Navigation { get; set; } = default!;

        [Inject]
        private IJSRuntime JS { get; set; } = default!;

        // Data Collections
        private List<BankBranchRecord> Branches { get; set; } = new();
        private List<BankAccountItem> Accounts { get; set; } = new();
        private List<string> DistinctBankNames { get; set; } = new();
        private List<string> DistinctCities { get; set; } = new();

        // Selected Entity State
        private string? SelectedBankName;
        private BankBranchRecord? SelectedBranch;
        private BankAccountItem? SelectedAccount;
        private List<BankAccountItem> SelectedBranchAccounts { get; set; } = new();

        // Tree Expansion & Search
        private string SearchText = string.Empty;
        private HashSet<string> ExpandedBanks = new(StringComparer.OrdinalIgnoreCase);
        private HashSet<int> ExpandedBranches = new();

        // Loading & Saving Spinners
        private bool IsLoading = true;
        private bool IsSaving = false;

        // Modal Controls & Form State
        private bool ShowBranchModal = false;
        private SaveBankBranchDto BranchModel = new();

        private bool ShowAccountModal = false;
        private SaveBankAccountDto AccountModel = new();

        private bool ShowODModal = false;
        private OverdraftConfigDto ODModel = new();

        private string strUserName = "System";
        private string strComputerName = "Web";

        // Hierarchical Grouping Model
        public class BankHierarchyGroup
        {
            public string BankName { get; set; } = string.Empty;
            public List<BranchHierarchyGroup> Branches { get; set; } = new();
        }

        public class BranchHierarchyGroup
        {
            public BankBranchRecord Branch { get; set; } = new();
            public List<BankAccountItem> Accounts { get; set; } = new();
        }

        protected override async Task OnInitializedAsync()
        {
            IsLoading = true;
            try
            {
                strUserName = _auditService.GetCurrentUserName();
                strComputerName = _auditService.GetClientIpAddress();

                await LoadDataAsync();

                // Auto-select Default Bank Account (or first account / branch if available)
                var defaultAcc = Accounts.FirstOrDefault(a => a.DefaultBank) ?? Accounts.FirstOrDefault();
                if (defaultAcc != null)
                {
                    SelectAccount(defaultAcc);
                }
                else if (Branches.Any())
                {
                    SelectBranch(Branches.First());
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Load Error", $"Failed to load banking data: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
            }
        }

        public void Dispose()
        {
        }

        private async Task LoadDataAsync()
        {
            Branches = await BankListDataAccess.GetAllBankBranchesAsync();
            Accounts = await BankListDataAccess.GetAllBankAccountsAsync();
            DistinctBankNames = await BankListDataAccess.GetDistinctBankNamesAsync();
            DistinctCities = await BankListDataAccess.GetDistinctCitiesAsync();

            // Auto-expand all banks and branches by default for instant visibility
            foreach (var b in Branches)
            {
                if (!string.IsNullOrWhiteSpace(b.Bank))
                {
                    ExpandedBanks.Add(b.Bank);
                }
                ExpandedBranches.Add(b.BankID);
            }

            // Sync currently selected account or branch if one was selected
            if (SelectedAccount != null)
            {
                SelectedAccount = Accounts.FirstOrDefault(a => a.AccNo == SelectedAccount.AccNo);
            }
            if (SelectedBranch != null)
            {
                SelectedBranch = Branches.FirstOrDefault(b => b.BankID == SelectedBranch.BankID);
                if (SelectedBranch != null)
                {
                    SelectedBranchAccounts = Accounts.Where(a => a.BankID == SelectedBranch.BankID).ToList();
                }
            }
        }

        // ─────────────────────────────────────────────────────────────
        // Tree Hierarchy Building & Filtering
        // ─────────────────────────────────────────────────────────────

        private List<BankHierarchyGroup> GetFilteredHierarchy()
        {
            var search = SearchText?.Trim() ?? string.Empty;
            var list = new List<BankHierarchyGroup>();

            // Distinct bank names from branches
            var bankNames = Branches
                .Select(b => b.Bank)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => name);

            foreach (var bankName in bankNames)
            {
                var branchList = Branches
                    .Where(b => string.Equals(b.Bank, bankName, StringComparison.OrdinalIgnoreCase))
                    .OrderBy(b => b.Branch)
                    .ToList();

                var branchHierarchyList = new List<BranchHierarchyGroup>();

                foreach (var branch in branchList)
                {
                    var branchAccounts = Accounts
                        .Where(a => a.BankID == branch.BankID)
                        .OrderByDescending(a => a.DefaultBank)
                        .ThenBy(a => a.AccTitle)
                        .ToList();

                    // Apply search filter if present
                    if (!string.IsNullOrWhiteSpace(search))
                    {
                        bool branchMatches =
                            branch.Bank.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                            branch.Branch.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                            branch.City.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                            (branch.Address != null && branch.Address.Contains(search, StringComparison.OrdinalIgnoreCase));

                        var matchingAccounts = branchAccounts.Where(a =>
                            a.AccTitle.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                            a.AccNo.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                            a.BankAccNo.Contains(search, StringComparison.OrdinalIgnoreCase)
                        ).ToList();

                        if (branchMatches || matchingAccounts.Any())
                        {
                            branchHierarchyList.Add(new BranchHierarchyGroup
                            {
                                Branch = branch,
                                Accounts = branchMatches ? branchAccounts : matchingAccounts
                            });
                        }
                    }
                    else
                    {
                        branchHierarchyList.Add(new BranchHierarchyGroup
                        {
                            Branch = branch,
                            Accounts = branchAccounts
                        });
                    }
                }

                if (branchHierarchyList.Any())
                {
                    list.Add(new BankHierarchyGroup
                    {
                        BankName = bankName,
                        Branches = branchHierarchyList
                    });
                }
            }

            return list;
        }

        // ─────────────────────────────────────────────────────────────
        // Tree Selection Handlers
        // ─────────────────────────────────────────────────────────────

        private void OnBankHeaderClick(BankHierarchyGroup bankGroup)
        {
            if (ExpandedBanks.Contains(bankGroup.BankName))
            {
                ExpandedBanks.Remove(bankGroup.BankName);
            }
            else
            {
                ExpandedBanks.Add(bankGroup.BankName);
            }

            SelectedBankName = bankGroup.BankName;
            var firstBranch = bankGroup.Branches.FirstOrDefault()?.Branch;
            if (firstBranch != null)
            {
                SelectBranch(firstBranch);
            }
        }

        private void OnBranchHeaderClick(BranchHierarchyGroup branchGroup)
        {
            if (ExpandedBranches.Contains(branchGroup.Branch.BankID))
            {
                ExpandedBranches.Remove(branchGroup.Branch.BankID);
            }
            else
            {
                ExpandedBranches.Add(branchGroup.Branch.BankID);
            }

            SelectBranch(branchGroup.Branch);
        }

        private void SelectBranch(BankBranchRecord branch)
        {
            SelectedBranch = branch;
            SelectedBankName = branch.Bank;
            SelectedBranchAccounts = Accounts.Where(a => a.BankID == branch.BankID).ToList();

            // If selected account doesn't belong to this branch, select first account in branch
            if (SelectedAccount == null || SelectedAccount.BankID != branch.BankID)
            {
                SelectedAccount = SelectedBranchAccounts.FirstOrDefault();
            }

            StateHasChanged();
        }

        private void SelectAccount(BankAccountItem acc)
        {
            SelectedAccount = acc;
            SelectedBranch = Branches.FirstOrDefault(b => b.BankID == acc.BankID);
            SelectedBankName = acc.Bank;

            if (SelectedBranch != null)
            {
                SelectedBranchAccounts = Accounts.Where(a => a.BankID == acc.BankID).ToList();
                ExpandedBranches.Add(SelectedBranch.BankID);
            }
            if (!string.IsNullOrWhiteSpace(acc.Bank))
            {
                ExpandedBanks.Add(acc.Bank);
            }

            StateHasChanged();
        }

        // ─────────────────────────────────────────────────────────────
        // Modal Handlers: Bank Branch
        // ─────────────────────────────────────────────────────────────

        private void OpenAddBranchModal()
        {
            BranchModel = new SaveBankBranchDto
            {
                BankID = 0,
                Bank = SelectedBankName ?? string.Empty,
                City = "Sialkot"
            };
            ShowBranchModal = true;
        }

        private void OpenAddBranchForBank(string? bankName)
        {
            BranchModel = new SaveBankBranchDto
            {
                BankID = 0,
                Bank = bankName ?? string.Empty,
                City = "Sialkot"
            };
            ShowBranchModal = true;
        }

        private void OpenEditBranchModal(BankBranchRecord branch)
        {
            BranchModel = new SaveBankBranchDto
            {
                BankID = branch.BankID,
                Bank = branch.Bank,
                Branch = branch.Branch,
                City = branch.City,
                Phone1 = branch.Phone1,
                Phone2 = branch.Phone2,
                Fax = branch.Fax,
                Email = branch.Email,
                Manager = branch.Manager,
                Address = branch.Address,
                Remittance = branch.Remittance,
                CompanyRefID = branch.CompanyRefID
            };
            ShowBranchModal = true;
        }

        private async Task SaveBranch()
        {
            if (string.IsNullOrWhiteSpace(BranchModel.Bank))
            {
                NotificationServiceManager.ShowWarning("Validation", "Bank Name is required.");
                return;
            }
            if (string.IsNullOrWhiteSpace(BranchModel.Branch))
            {
                NotificationServiceManager.ShowWarning("Validation", "Branch Name is required.");
                return;
            }

            IsSaving = true;
            try
            {
                var result = await BankListDataAccess.SaveBankBranchAsync(BranchModel, strUserName, strComputerName);
                if (result.Success)
                {
                    NotificationServiceManager.ShowSuccess("Success", result.Message);
                    ShowBranchModal = false;
                    await LoadDataAsync();

                    var savedBranch = Branches.FirstOrDefault(b => b.BankID == result.BankID);
                    if (savedBranch != null)
                    {
                        SelectBranch(savedBranch);
                    }
                }
                else
                {
                    NotificationServiceManager.ShowError("Save Failed", result.Message);
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error", ex.Message);
            }
            finally
            {
                IsSaving = false;
            }
        }

        // ─────────────────────────────────────────────────────────────
        // Modal Handlers: Bank Account
        // ─────────────────────────────────────────────────────────────

        private void OpenAddAccountModal(int bankId)
        {
            var branch = Branches.FirstOrDefault(b => b.BankID == bankId);
            string defaultTitle = branch != null ? $"{branch.Bank} - {branch.Branch}" : string.Empty;

            AccountModel = new SaveBankAccountDto
            {
                BankID = bankId,
                IsNew = true,
                Type = "Current",
                AccTitle = defaultTitle,
                OpenBal = 0,
                OpenDate = DateTime.Today,
                DefaultBank = !Accounts.Any() // automatically make default if it's the very first account
            };
            ShowAccountModal = true;
        }

        private void OpenEditAccountModal(BankAccountItem acc)
        {
            AccountModel = new SaveBankAccountDto
            {
                BankID = acc.BankID,
                AccNo = acc.AccNo,
                AccTitle = acc.AccTitle,
                BankAccNo = acc.BankAccNo,
                Type = acc.Type,
                DefaultBank = acc.DefaultBank,
                AllowOD = acc.AllowOD,
                ODAmt = acc.ODAmt,
                ODRate = acc.ODRate,
                Chqformat = acc.Chqformat,
                OpenBal = acc.OpenBal,
                OpenDate = acc.OpenDate,
                IsNew = false
            };
            ShowAccountModal = true;
        }

        private async Task SaveAccount()
        {
            if (AccountModel.BankID <= 0)
            {
                NotificationServiceManager.ShowWarning("Validation", "Please select a target Bank Branch.");
                return;
            }
            if (string.IsNullOrWhiteSpace(AccountModel.AccTitle))
            {
                NotificationServiceManager.ShowWarning("Validation", "Account Title is required.");
                return;
            }
            if (string.IsNullOrWhiteSpace(AccountModel.BankAccNo))
            {
                NotificationServiceManager.ShowWarning("Validation", "Bank Account Number is required.");
                return;
            }

            IsSaving = true;
            try
            {
                var result = await BankListDataAccess.SaveBankAccountAsync(AccountModel, strUserName, strComputerName);
                if (result.Success)
                {
                    NotificationServiceManager.ShowSuccess("Success", $"{result.Message} (GL: {result.GeneratedAccNo})");
                    ShowAccountModal = false;
                    await LoadDataAsync();

                    var savedAcc = Accounts.FirstOrDefault(a => a.AccNo == result.GeneratedAccNo);
                    if (savedAcc != null)
                    {
                        SelectAccount(savedAcc);
                    }
                }
                else
                {
                    NotificationServiceManager.ShowError("Save Failed", result.Message);
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error", ex.Message);
            }
            finally
            {
                IsSaving = false;
            }
        }

        // ─────────────────────────────────────────────────────────────
        // Overdraft & Default Bank Operations
        // ─────────────────────────────────────────────────────────────

        private async Task SetAsDefault(string accNo)
        {
            try
            {
                var result = await BankListDataAccess.SetDefaultBankAccountAsync(accNo);
                if (result.Success)
                {
                    NotificationServiceManager.ShowSuccess("Default Bank Set", result.Message);
                    await LoadDataAsync();
                }
                else
                {
                    NotificationServiceManager.ShowError("Operation Failed", result.Message);
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error", ex.Message);
            }
        }

        private void OpenODModal(BankAccountItem acc)
        {
            ODModel = new OverdraftConfigDto
            {
                AccNo = acc.AccNo,
                AllowOD = acc.AllowOD,
                ODAmt = acc.ODAmt ?? 0,
                ODRate = acc.ODRate ?? 0
            };
            ShowODModal = true;
        }

        private async Task SaveOD()
        {
            IsSaving = true;
            try
            {
                var result = await BankListDataAccess.UpdateOverdraftAsync(ODModel);
                if (result.Success)
                {
                    NotificationServiceManager.ShowSuccess("Overdraft Updated", result.Message);
                    ShowODModal = false;
                    await LoadDataAsync();
                }
                else
                {
                    NotificationServiceManager.ShowError("Failed", result.Message);
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error", ex.Message);
            }
            finally
            {
                IsSaving = false;
            }
        }

        // ─────────────────────────────────────────────────────────────
        // Delete Confirmations
        // ─────────────────────────────────────────────────────────────

        private async Task OpenDeleteAccountConfirm(BankAccountItem acc)
        {
            bool? confirmed = await DialogService.Confirm(
                $"Are you sure you want to permanently delete Bank Account '[{acc.AccNo}] {acc.AccTitle}'?\n\nThis will remove the mapping from BankAccounts and Accounts.",
                "Delete Bank Account",
                new ConfirmOptions
                {
                    OkButtonText = "Yes, Delete",
                    CancelButtonText = "Cancel",
                    AutoFocusFirstElement = false
                });

            if (confirmed == true)
            {
                try
                {
                    var result = await BankListDataAccess.DeleteBankAccountAsync(acc.AccNo);
                    if (result.Success)
                    {
                        NotificationServiceManager.ShowSuccess("Deleted", result.Message);
                        SelectedAccount = null;
                        await LoadDataAsync();
                    }
                    else
                    {
                        NotificationServiceManager.ShowWarning("Cannot Delete", result.Message);
                    }
                }
                catch (Exception ex)
                {
                    NotificationServiceManager.ShowError("Error", ex.Message);
                }
            }
        }

        private async Task OpenDeleteBranchConfirm(BankBranchRecord branch)
        {
            bool? confirmed = await DialogService.Confirm(
                $"Are you sure you want to delete branch '{branch.Bank} — {branch.Branch}'?",
                "Delete Bank Branch",
                new ConfirmOptions
                {
                    OkButtonText = "Yes, Delete",
                    CancelButtonText = "Cancel",
                    AutoFocusFirstElement = false
                });

            if (confirmed == true)
            {
                try
                {
                    var result = await BankListDataAccess.DeleteBankBranchAsync(branch.BankID);
                    if (result.Success)
                    {
                        NotificationServiceManager.ShowSuccess("Deleted", result.Message);
                        SelectedBranch = null;
                        SelectedAccount = null;
                        await LoadDataAsync();
                    }
                    else
                    {
                        NotificationServiceManager.ShowWarning("Cannot Delete", result.Message);
                    }
                }
                catch (Exception ex)
                {
                    NotificationServiceManager.ShowError("Error", ex.Message);
                }
            }
        }

        // ─────────────────────────────────────────────────────────────
        // Navigation Handlers
        // ─────────────────────────────────────────────────────────────

        private void GoToChequeBooks()
        {
            Navigation.NavigateTo("/accounts/chqbookdetail");
        }

        private void GoBack()
        {
            if (!string.IsNullOrEmpty(ReturnUrl))
            {
                Navigation.NavigateTo(ReturnUrl);
            }
            else
            {
                Navigation.NavigateTo("/");
            }
        }
    }
}