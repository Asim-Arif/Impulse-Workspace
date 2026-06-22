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
using Microsoft.AspNetCore.Components.Web;
using Microsoft.CodeAnalysis.Text;
using Microsoft.JSInterop;
using Radzen;

namespace Impulse.Pages.Accounts
{
    public partial class BankBalanceStatement : ComponentBase
    {
        private string? EditingAccNo { get; set; }
        private bool _shouldFocus;
        private string _focusId;

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

        private List<Bank_Balance_Statement_ViewModel> ListToUse { get; set; }=new List<Bank_Balance_Statement_ViewModel>();
        private Trial_Balance_ViewModel SelectedRow =new Trial_Balance_ViewModel();
        //If it feels slow, could use filtered list to only filter once for all 6 totals like 
        //var filteredList = ListToUse.Where(x => x.EntryType == 0).ToList();
        //and then use filteredList to Sum
        private decimal TotalBankBalance => ListToUse?.Sum(x => x.Balance) ?? 0;
        
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
                await RefreshList();

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
        protected override async Task OnAfterRenderAsync(bool firstRender)
        {
            if (_shouldFocus)
            {
                _shouldFocus = false;
                //await JS.InvokeVoidAsync("focusElement", _focusId);
            }
        }
        private void GoToIndexPage()
        {
            Navigation.NavigateTo("/", true);
        }        
        private async Task RefreshList()
        {
            IsLoading = true;
            try
            {                
                ListToUse.Clear();

                //var bodList = await AccountReportingAccess.GetLedgerData(CurrentAccount.AccNo,CurrentAccount.DTFrom,CurrentAccount.DTTo);
                var listtouse = await AccountReportingAccessService.GetBankBalanceStatement();
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

        private async Task StartEdit(Bank_Balance_Statement_ViewModel row)
        {
            EditingAccNo = row.AccNo;
            //await JS.InvokeVoidAsync("focusElement", $"input-{row.AccNo}");
            _focusId = $"input-{row.AccNo}";
            _shouldFocus = true;
        }

        private async Task SaveEdit(Bank_Balance_Statement_ViewModel row)
        {
            if (!string.IsNullOrEmpty(EditingAccNo))
            {

                await IDBHelperService.ExecuteAsync("UPDATE BankAccounts SET ChqFormat=@ChqFormat WHERE AccNo=@AccNo",
                    new 
                    {
                        ChqFormat=row.ChqFormat,
                        AccNo=row.AccNo
                    });
                EditingAccNo = null;
                StateHasChanged(); // Use this sparingly inside async interop calls
            }
        }

        private async Task HandleKeyDown(KeyboardEventArgs e, Bank_Balance_Statement_ViewModel row)
        {
            if (e.Key == "Enter")
            {
                await SaveEdit(row);
            }
            else if (e.Key == "Escape")
            {
                EditingAccNo = null;                
            }
        }
        private void PrintBankBalanceSheet() 
        {                        
            var reportRequest = new ReportRequest
            {
                ReportName = ReportNames.CashnBankStatus,
                SelectionFormula = $"1=1"                
            };

            ReportNavigationService.PrintReportAsync(reportRequest);                        
        }
    }
}