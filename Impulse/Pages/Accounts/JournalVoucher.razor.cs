using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Impulse.Components.Common;
using Impulse.Services;
using Impulse.Services.Common;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Radzen;

namespace Impulse.Pages.Accounts
{
    public partial class JournalVoucher : ComponentBase, IDisposable
    {
        private string strComputerName = string.Empty;
        private string strUserName = string.Empty;
        private string strCashInHandParent=string.Empty;
        private VoucherViewModel voucherViewModel = new VoucherViewModel();
        private string VoucherNo { get; set; }
        private string successMessage = string.Empty;
        private bool isError = false;
        private bool isSaving = false;
        private AttachmentModal? attachmentModalRef;
        private int attachmentCount = 0;
        private DateTime selectedDate = DateTime.Today;
        private string strCashBalance { get; set; } = string.Empty;
        private string strAccountBalance { get; set; } = string.Empty;
        private decimal dDebitCreditDifference { get; set; } = 0;
        
        private List<ChartOfAccountsModel> Accounts = new List<ChartOfAccountsModel>();
        
        private ChartOfAccountsModel? SelectedAccount = null;        

        private VoucherLineItemViewModel newVoucherLine = new VoucherLineItemViewModel();
        private string validationMessage = string.Empty;
        private bool IsPrintWhenSaved = true;

        private Timer _timer; // Declare the Timer object
        // Implement IDisposable to ensure the timer is cleaned up
        public void Dispose()
        {
            _timer?.Dispose();
        }
        protected override async Task OnInitializedAsync()
        {
            await InitializeData();

            strComputerName= _auditService.GetClientIpAddress();
            strUserName=  _auditService.GetCurrentUserName();

            strCashInHandParent = await _idbHelperService.getSingleStringValue("DataValue", "GeneralData", "WHERE DataName='CashInHandParent'");

            Accounts = await _voucherService.GetTransactionalAccounts("WHERE SubAccOf<>'"+strCashInHandParent+"'");
            string strHODCondition = " WHERE (UserName<>'" + strUserName + "' AND InActive=0 AND UserRole NOT IN(''))";

        }
        // This method handles the timer expiration
        private void HideSuccessMessage(object state)
        {
            // Use InvokeAsync to ensure the UI update runs on Blazor's synchronization context
            InvokeAsync(() =>
            {
                successMessage = string.Empty;
                StateHasChanged();
            });
            _timer?.Dispose();
        }

        [Inject]
        private IVoucherService _voucherService { get; set; }
        [Inject]
        private IAuditService _auditService { get; set; }
        
        [Inject]
        private IDBHelperService _idbHelperService { get; set; }

        [Inject]
        private INotificationService NotificationServiceManager { get; set; }

        [Inject]
        private IReportNavigationService ReportNavigationService { get; set; }

        [Inject]
        private IEntityAttachmentService AttachmentService { get; set; } = default!;

        [Parameter]
        public int p_bankId { get; set; }


        private async Task GetNextVoucherNo(DateTime selectedDate)
        {
            voucherViewModel.VchrNo = await VoucherService.GetNextVchrNo(selectedDate,"JV");
            await LoadAttachmentCountAsync();
        }

        private async Task OnDTChanged(ChangeEventArgs e)
        {
            await GetNextVoucherNo(voucherViewModel.DT);
            StateHasChanged();
        }

        private async Task InitializeData()
        {
            isSaving = false;

            /*
            if (voucherViewModel.ChequeDetails == null)
            {
                voucherViewModel.ChequeDetails = new ChequeViewModel();
            }
            voucherViewModel.ChequeDetails.ChequeDate = DateTime.Today;
            voucherViewModel.ChequeDetails.ChequeNo = string.Empty;
            voucherViewModel.ChequeDetails.Amount = 0;
            */

            voucherViewModel.DT = DateTime.Today;
            
            //strCashBalance = string.Empty;
            voucherViewModel.LineItems.Clear();
            voucherViewModel.Notes = string.Empty;
            dDebitCreditDifference = 0;

            
            // Later i'll change amount & get from GD         
            await GetNextVoucherNo(voucherViewModel.DT);
            StateHasChanged();

        }

        //private List<VoucherLineItemViewModel> VoucherLineList = new List<VoucherLineItemViewModel>();
        
        
        
        private async Task AddEntry()
        {

            if (SelectedAccount == null)
            {
                //validationMessage = "Please Select Account";
                NotificationServiceManager.ShowError("Please select Account No.", "You need to select account from the drop-down.");
                return;
            }
            if (string.IsNullOrEmpty(newVoucherLine.Description)) 
            {
                NotificationServiceManager.ShowError("Error", "Please enter Description.");
                return;
            }
            if (newVoucherLine.Debit == 0 && newVoucherLine.Credit == 0) 
            {
                NotificationServiceManager.ShowError("Error", "Debit and Credit cannot both be zero.");
                return;
            }
            if (newVoucherLine.Debit != 0 && newVoucherLine.Credit != 0)
            {
                NotificationServiceManager.ShowError("Error", "Only one field (Debit or Credit) can have a value.");
                return;
            }
            if (newVoucherLine.Debit < 0 || newVoucherLine.Credit < 0)
            {
                NotificationServiceManager.ShowError("Error", "You can't have negative value in field (Debit or Credit).");
                return;
            }
            
            var lineToAdd = new VoucherLineItemViewModel
            {
                AccNo = SelectedAccount.AccNo,
                AccTitle = SelectedAccount.AccTitle,
                Description = newVoucherLine.Description,
                Debit = newVoucherLine.Debit,
                Credit = newVoucherLine.Credit
            };

            voucherViewModel.LineItems.Add(lineToAdd);
            
            //Update Cheque Amount
            //voucherViewModel.ChequeDetails.Amount = voucherViewModel.TotalDebit - voucherViewModel.TotalCredit;
            dDebitCreditDifference = voucherViewModel.TotalDebit - voucherViewModel.TotalCredit;
            // Clear form inputs
            newVoucherLine.Description = string.Empty;
            newVoucherLine.Debit = 0;
            newVoucherLine.Credit = 0;

            // flight no should not b blank after adding entry into list
            //SelectedFlightNo.FlightNo = string.Empty;
            //SelectedAccount.AccTitle = string.Empty;
            SelectedAccount = null;
            // Clear the validation message
            validationMessage = string.Empty;

            try
            {
                _ = JS.InvokeVoidAsync("eval", "setTimeout(() => { document.querySelector('.enter-as-tab .blazored-typeahead__input')?.focus(); }, 60);");
            }
            catch { }
        }

        private void RemoveVoucherLine(VoucherLineItemViewModel VLIVM)
        {
            voucherViewModel.LineItems.Remove(VLIVM);
            //voucherViewModel.ChequeDetails.Amount = voucherViewModel.TotalDebit - voucherViewModel.TotalCredit;
            //VoucherLineList.Remove(VLIVM);
        }
        // NOW DATE DATA INTO DATABSE 


        private async Task SaveVoucherMain()
        {
            if (isSaving)
                return;

            VoucherLineItemViewModel systemLine = null;
            try
            {
                isSaving = true;
                // Ensure cipList has items
                if (voucherViewModel.LineItems == null || !voucherViewModel.LineItems.Any())
                {
                    NotificationServiceManager.ShowWarning("Error", "No items added. Please add items before saving.");
                    return;
                }                
                if (string.IsNullOrEmpty(voucherViewModel.Notes))
                {
                    NotificationServiceManager.ShowWarning("Error", "Please enter Notes.");
                    return;
                }
                decimal dTotalDebit = 0;
                decimal dTotalCredit = 0;

                dTotalDebit=voucherViewModel.TotalDebit;
                dTotalCredit = voucherViewModel.TotalCredit;
                if (dTotalDebit != dTotalCredit) 
                {
                    NotificationServiceManager.ShowWarning("Error", "Can't post transaction when Debit and Credit are not equal.");
                    return;
                }
                
                

                string strReturnVchrNo = await VoucherService.SaveVoucherAsync(voucherViewModel);

                isError = false;
                //successMessage = "Data Saved Successfully!";
                NotificationServiceManager.ShowSuccess("Saved", "Transaction saved successfully.");
                
                if (IsPrintWhenSaved)
                {
                    ReportNavigationService.PrintVoucher(strReturnVchrNo);
                }
                //StateHasChanged();
                // Auto-hide the message after 10 seconds
                //await Task.Delay(3000);
                //_timer?.Dispose();
                // 2. Initialize a new timer that calls HideSuccessMessage after 3000ms (3 seconds)
                // Note: The timer starts immediately.
                //_timer = new Timer(
                //    callback: HideSuccessMessage,
                //    state: null,
                //    dueTime: 30000, // Wait 3000 milliseconds before first call
                //    period: Timeout.Infinite // Do not repeat
                //);

                //successMessage = string.Empty;
                //StateHasChanged(); // Update UI

                await ResetForm();
                
                //VoucherLineList.Clear();
                isSaving = false;
            }
            catch (Exception ex)
            {
                isError = true;
                //successMessage = $"{ex.Message}";
                /*_timer = new Timer(
                    callback: HideSuccessMessage,
                    state: null,
                    dueTime: 10000, // Maybe keep error messages visible longer (e.g., 10 seconds)
                    period: Timeout.Infinite
                );*/
                NotificationServiceManager.ShowError("Error", $"{ex.Message}");
                if (systemLine != null)
                {
                    voucherViewModel.LineItems.Remove(systemLine);
                }
                isSaving = false;
            }
            finally
            {
                isSaving = false;
            }
        }

        // RESET FORM AFTER SAVE DATA INTO DATABASE
        private async Task ResetForm()
        {
            await InitializeData();
            StateHasChanged();
        }

        private void GoToIndexPage()
        {
            Navigation.NavigateTo("/financial");
        }

        private void OpenAttachmentModal()
        {
            if (string.IsNullOrWhiteSpace(voucherViewModel.VchrNo)) return;

            attachmentModalRef?.Open(
                entityType: "Voucher",
                entityId: voucherViewModel.VchrNo,
                displayName: $"Journal Voucher #{voucherViewModel.VchrNo}"
            );
        }

        private void OnAttachmentCountChanged(int count)
        {
            attachmentCount = count;
            StateHasChanged();
        }

        private async Task LoadAttachmentCountAsync()
        {
            if (!string.IsNullOrWhiteSpace(voucherViewModel.VchrNo))
            {
                try
                {
                    attachmentCount = await AttachmentService.GetCountAsync("Voucher", voucherViewModel.VchrNo);
                }
                catch
                {
                    attachmentCount = 0;
                }
            }
            else
            {
                attachmentCount = 0;
            }
        }
        
        
        private async Task<IEnumerable<ChartOfAccountsModel>> GetTransactionAccounts(string searchText)
        {
            return await Task.FromResult(Accounts.Where(x => x.AccTitle.ToLower().Contains(searchText.ToLower())).ToList());
        }

        private async Task SelectedResultChanged_Account(ChartOfAccountsModel? selectedAccount)
        {
            SelectedAccount = selectedAccount;
            if (selectedAccount != null && selectedAccount.AccTitle != null)
            {
                newVoucherLine.AccTitle = selectedAccount.AccTitle;
                newVoucherLine.AccNo = selectedAccount.AccNo;
                
                decimal balance = await VoucherService.GetBalance(selectedAccount.AccNo, voucherViewModel.DT);
                strAccountBalance = balance.ToString("N2");
                StateHasChanged();
                
            }
        }

    }
}
