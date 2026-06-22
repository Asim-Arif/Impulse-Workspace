using BlazorBootstrap;
using BlazorContextMenu;
using Impulse.Services.Integrations;
using DataAccessLibrary;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Humanizer;
using Impulse.Constants;
using Impulse.Pages.SelectionModals;
using Impulse.Services;
using Impulse.Services.Integrations;
using Microsoft.AspNetCore.Components;
using Microsoft.CodeAnalysis.Text;
using Microsoft.JSInterop;
using Radzen;

namespace Impulse.Pages.Accounts
{
    public partial class TrialBalance : ComponentBase
    {
        [Parameter] public int? TrialType { get; set; }
        private bool IsDropdownDisabled { get; set; } = false; // Default to enabled
        protected GetSingleDropdownValue ChequeBookModal;
        // List to hold the cheque book dropdown items
        protected List<GenericDropDownModel> ChequeBookList = new();

        private bool IsLoading=true;
        [Inject]
        private IVoucherService IVoucherService { get; set; }
        [Inject]
        private IDBHelperService IDBHelperService { get; set; }
        [Inject]
        protected IAccountReportingAccessService AccountReportingAccessService { get; set; }

        [Inject]
        protected IReportNavigationService ReportNavigationService { get; set; }
        [Inject]
        private INotificationService NotificationServiceManager { get; set; } = null!;
        [Inject]
        private IAuditService AuditService { get; set; }

        [Inject]
        private IMeezanBankService meezanBankService { get; set; }

        [Inject]
        private IFaysalBankServices faysalBankServices { get; set; }
        
        private List<GenericDropDownModel> AccountHeads = new List<GenericDropDownModel>();
        private GenericDropDownModel? SelectedAccount = null;

        private List<Trial_Balance_ViewModel> ListToUse { get; set; }=new List<Trial_Balance_ViewModel>();
        private Trial_Balance_ViewModel SelectedRow =new Trial_Balance_ViewModel();
        //If it feels slow, could use filtered list to only filter once for all 6 totals like 
        //var filteredList = ListToUse.Where(x => x.EntryType == 0).ToList();
        //and then use filteredList to Sum
        private decimal TotalOpeningDr => ListToUse?.Where(x => x.EntryType == 0).Sum(x => x.Opening_Debit) ?? 0;
        private decimal TotalOpeningCr => ListToUse?.Where(x => x.EntryType == 0).Sum(x => x.Opening_Credit) ?? 0;
        private decimal TotalDebit => ListToUse?.Where(x => x.EntryType == 0).Sum(x => x.Debit) ?? 0;
        private decimal TotalCredit => ListToUse?.Where(x => x.EntryType == 0).Sum(x => x.Credit) ?? 0;
        private decimal TotalBalanceDr => ListToUse?.Where(x => x.EntryType == 0).Sum(x => x.Balance_Debit) ?? 0;
        private decimal TotalBalanceCr => ListToUse?.Where(x => x.EntryType == 0).Sum(x => x.Balance_Credit) ?? 0;
        private DateTime DTFrom { get; set; } = DateTime.Now;
        private DateTime DTTo { get; set; } = DateTime.Now;

        private int SelectedStatus { get; set; } = 0;
        private string SelectedRangeValue { get; set; } = "0";
        protected override async Task OnInitializedAsync()
        {
            IsLoading = true;
            try
            {
                IsLoading = true;
                AccountHeads = await IVoucherService.GetValuesForDropDown("Accounts", "AccNo", "AccTitle"," ORDER BY AccNo");
                
                AccountHeads.Insert(0, new GenericDropDownModel()
                {
                    DropDownValue_ID="0",
                    DropDownValue_Description="<All Heads>"
                });

                if (TrialType.HasValue)
                {
                    if (TrialType == 1) //Suppliers
                    {
                        IsDropdownDisabled = true;
                        string strAccountTitle = await IDBHelperService.getGeneralDataValueAsync("TB_AdvanceToSupplier");
                        var match = AccountHeads.FirstOrDefault(p => p.DropDownValue_Description == strAccountTitle);
                        if (match != null)
                        {
                            SelectedAccount = match; //AccountHeads.FirstOrDefault(p => p.DropDownValue_Description == strAccountTitle) ?? SelectedAccount;
                            await SelectedResultChanged_AccountHead(SelectedAccount);
                        }                                                
                    }
                    
                }
                else if (AccountHeads != null && AccountHeads.Any()) //Select Default 
                {
                    SelectedAccount = AccountHeads.FirstOrDefault(p => p.DropDownValue_ID == "0") ?? AccountHeads.First();
                    await SelectedResultChanged_AccountHead(SelectedAccount);
                }

                if (DateTime.Today.Month >= 7)
                {
                    DTFrom = new DateTime(DateTime.Today.Year, 7, 1);
                    DTTo = new DateTime(DateTime.Today.Year + 1, 6, 30);
                }
                else                
                {
                    DTFrom = new DateTime(DateTime.Today.Year - 1, 7, 1);
                    DTTo= new DateTime(DateTime.Today.Year , 6, 30);
                }

                IsLoading = false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching Data: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
            }
        }
        private void GoToIndexPage()
        {
            Navigation.NavigateTo("/", true);
        }

        
        private async Task<IEnumerable<GenericDropDownModel>> AccountHeadSearch(string searchText)
        {
            return await Task.FromResult(
                AccountHeads
                    .Where(x => x.DropDownValue_Description
                        .Contains(searchText,StringComparison.OrdinalIgnoreCase))
                    .ToList());
        }
        
        private async Task SelectedResultChanged_AccountHead(GenericDropDownModel? selectedValue)
        {
            SelectedAccount = selectedValue;
        }
        
        private async Task RefreshList()
        {
            IsLoading = true;
            try
            {                
                ListToUse.Clear();

                //var bodList = await AccountReportingAccess.GetLedgerData(CurrentAccount.AccNo,CurrentAccount.DTFrom,CurrentAccount.DTTo);
                var listtouse = await AccountReportingAccessService.GetTrialBalanceList(DTFrom,DTTo, SelectedAccount.DropDownValue_ID);
                ListToUse = listtouse.ToList();                
                

                StateHasChanged();

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching Data: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
                StateHasChanged();
            }
        }

        
        private void PrintVoucher(ItemClickEventArgs e)
        {
            /*
            var vchr = e.Data as AccountsReportingModel;
            CurrentAccount = vchr;

            if (CurrentAccount != null)
            {
                ReportNavigationService.PrintVoucher(CurrentAccount.VchrNo);                
            }
            */
        }
        private void GoToVoucherTransaction() 
        {
        }
        private void PrintTrialBalanceReport() 
        {
            string strFromTo = DTFrom.ToString("dd-MMM-yyyy") + " to " + DTTo.ToString("dd-MMM-yyyy");            
            
            var reportRequest = new ReportRequest
            {
                ReportName = ReportNames.TrialBalance,
                SelectionFormula = $"1=1",
                FormulaValues = new Dictionary<string, object>
                {
                    { "FromTo", $"\"{strFromTo}\"" }
                },
                Parameters = new Dictionary<string, object>
                {
                    { "@DTFrom", DTFrom },
                    { "@DTTo", DTTo },
                    { "@Code", SelectedAccount.DropDownValue_ID  },
                }
            };

            ReportNavigationService.PrintReportAsync(reportRequest);                        
        }
    }
}