using BlazorBootstrap;
using DataAccessLibrary;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Microsoft.AspNetCore.Components;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Accounts
{
    public partial class AccountsPayableReceivable : ComponentBase
    {
        [Parameter]
        public string Type { get; set; }

        [Inject]
        protected IChartOfAccountsDataAccess ChartOfAccountsService { get; set; }
        
        [Inject]
        protected NavigationManager Navigation { get; set; }

        [Inject]
        protected IDBHelper IDBHelper { get; set; }
        
        [Inject]
        private Impulse.Services.INotificationService NotificationServiceManager { get; set; }
        
        [Inject]
        protected Impulse.Services.IReportNavigationService ReportNavigationService { get; set; }

        private List<ChartOfAccountsModel> AccountsList { get; set; } = new List<ChartOfAccountsModel>();
        private ChartOfAccountsModel SelectedAccount { get; set; }
        private bool IsLoading;
        private bool bshowinactive = false;
        private string PageTitle => Type?.ToLower() == "receivable" ? "Accounts Receivable" : "Accounts Payable";
        private string RootAccNo => Type?.ToLower() == "receivable" ? "15-003" : "24-004";

        protected override async Task OnInitializedAsync()
        {
            await LoadData();
        }

        protected override async Task OnParametersSetAsync()
        {
            await LoadData();
        }

        private async Task LoadData()
        {
            IsLoading = true;
            StateHasChanged();
            
            try
            {
                var accountsFromDb = await ChartOfAccountsService.GetChartOfAccountsList(bshowinactive, RootAccNo);
                AccountsList = ProcessAccountsWithHeads(accountsFromDb);
            }
            catch (System.Exception ex)
            {
                System.Console.WriteLine($"Error fetching Data: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
                StateHasChanged();
            }
        }

        private List<ChartOfAccountsModel> ProcessAccountsWithHeads(IEnumerable<ChartOfAccountsModel> accounts)
        {
            var result = new List<ChartOfAccountsModel>();
            var grouped = accounts.GroupBy(a => new { a.HeadTypeNo, a.HeadType }).OrderBy(g => g.Key.HeadTypeNo);

            foreach (var group in grouped)
            {
                var groupList = group.OrderBy(a => a.AccNo).ToList();

                // Update Parent accounts to hold the sum of their children's balances.
                foreach (var acc in groupList)
                {
                    if (acc.AccType) // AccType == true means it's a Parent node
                    {
                        var children = groupList.Where(c => c.AccNo.StartsWith(acc.AccNo) && c.AccNo != acc.AccNo && !c.AccType);
                        acc.OpeningBalance = children.Sum(c => c.OpeningBalance);
                        acc.Balance = children.Sum(c => c.Balance);
                    }
                }

                // Add Head row
                result.Add(new ChartOfAccountsModel
                {
                    AccNo = group.Key.HeadTypeNo,
                    AccTitle = group.Key.HeadType,
                    bMainhead = false, // Renders as table-secondary fw-bold in the grid
                    AccType = false,
                    Active = true,
                    OpeningBalance = groupList.Where(a => !a.AccType).Sum(a => a.OpeningBalance), // Only sum leaf nodes to avoid double counting
                    Balance = groupList.Where(a => !a.AccType).Sum(a => a.Balance)
                });

                // Add accounts under this Head
                result.AddRange(groupList);
            }

            return result;
        }

        private async void RefreshData(BlazorContextMenu.ItemClickEventArgs e)
        {
            await LoadData();
        }

        private async Task ShowInActive(BlazorContextMenu.ItemClickEventArgs e, bool showinactive)
        {
            bshowinactive = showinactive;
            await LoadData();
        }

        private async Task MakeAccountActive(BlazorContextMenu.ItemClickEventArgs e)
        {
            try
            {
                var account = e.Data as ChartOfAccountsModel;
                SelectedAccount = account;

                if (SelectedAccount == null)
                    return;

                int isExitInTransaction = 0;
                bool bactive = SelectedAccount.Active;
                if (bactive == true)
                {
                    int bParentAccount = await IDBHelper.getSingleIntValueasync("Parent", "Accounts", $"WHERE AccNo= '{SelectedAccount.AccNo}'");
                    isExitInTransaction = await IDBHelper.getSingleIntValueasync("COUNT(*)", "FixedAccounts", $"WHERE AccNo = '{SelectedAccount.AccNo}'");

                    if (isExitInTransaction > 0)
                    {
                        NotificationServiceManager.ShowWarning("Cannot Make This Account Inactive.", "This Account Has A Fixed Account Number.");
                        return;
                    }

                    if (bParentAccount == 1)
                    {
                        int isChildAccountsExits = await IDBHelper.getSingleIntValueasync("COUNT(*)", "Accounts", $"WHERE SubAccOf = '{SelectedAccount.AccNo}' AND Active = 1");
                        if (isChildAccountsExits > 0)
                        {
                            NotificationServiceManager.ShowWarning("Cannot Make This Account Inactive.", "This Account Has Active Child Accounts. Please make them inactive first.");
                            return;
                        }
                    }
                    int isExitInTransaction_Vouchers = await IDBHelper.getSingleIntValueasync("COUNT(*)", "Vouchers", $"WHERE AccNo = '{SelectedAccount.AccNo}'");
                    if (isExitInTransaction_Vouchers > 0)
                    {
                        NotificationServiceManager.ShowWarning("Cannot Make InActive.", "There is Some Transaction Related To This Account.");
                        return;
                    }

                    await ChartOfAccountsService.UpdateAccountStatusAsync(SelectedAccount.AccNo, SelectedAccount.Active);
                }
                else
                {
                    await ChartOfAccountsService.UpdateAccountStatusAsync(SelectedAccount.AccNo, SelectedAccount.Active);
                }

                await LoadData();
            }
            catch (System.Exception ex)
            {
                NotificationServiceManager.ShowError("Error Occurred", $"Failed To Change Account Status: {ex.Message}");
            }
        }

        private void PrintChartOfAccounts(int itype)
        {
            string strcond = "";
            if (bshowinactive)
            {
                strcond = "1=1";   // show all records
            }
            else
            {
                strcond = "{VTempAccounts.Active} = 1";
            }

            // Since it's Payable/Receivable, also restrict by SubAccOf in formula
            strcond += $" AND ({{VTempAccounts.AccNo}} = '{RootAccNo}' OR {{VTempAccounts.SubAccOf}} startswith '{RootAccNo}')";

            string strreportname = "";
            if (itype == 1)
            {
                strreportname = Impulse.Constants.ReportNames.Accounts.PrintAccounts;
            }
            else
            {
                strreportname = Impulse.Constants.ReportNames.Accounts.AccountsWithOpening;
            }

            var reportRequest = new Impulse.Services.ReportRequest
            {
                ReportName = strreportname,
                SelectionFormula = strcond,
            };
            ReportNavigationService.PrintReportAsync(reportRequest);
        }

        private void GoToIndexPage()
        {
            Navigation.NavigateTo("/financial", true);
        }
    }
}
