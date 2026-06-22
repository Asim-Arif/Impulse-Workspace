using CIP.Services;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Forms;
using Radzen;

namespace CIP.Pages.Accounts
{
    public partial class CashReceiptVoucher : ComponentBase, IDisposable
    {
        
        private string strComputerName = string.Empty;
        private string strUserName = string.Empty;
        private string strCashInHandParent=string.Empty;
        private VoucherViewModel voucherViewModel = new VoucherViewModel();
        private float fTaxRate=0;
        private float fTaxAmount = 0;
        private string VoucherNo { get; set; }
        private bool bClearInputFile = false;        
        private bool EnableCostCenter = false;
        private string successMessage = string.Empty;
        private bool isError = false;
        private bool showTaxModal = false;
        private string myFileExtension = string.Empty;
        private bool isSaving = false;
        private DateTime selectedDate = DateTime.Today;
        private string strCashBalance { get; set; } = string.Empty;
        private string strAccountBalance { get; set; } = string.Empty;
        private decimal dDebitCreditDifference { get; set; } = 0;
        
        private List<ChartOfAccountsModel> cashAccounts = new List<ChartOfAccountsModel>();
        private List<ChartOfAccountsModel> Accounts = new List<ChartOfAccountsModel>();
        private List<GenericDropDownModel> CostCenters = new List<GenericDropDownModel>();
        private List<GenericDropDownModel> TaxHeadTypes = new List<GenericDropDownModel>();
        
        private ChartOfAccountsModel? SelectedCashAccount = null;
        private ChartOfAccountsModel? SelectedAccount = null;        
        private GenericDropDownModel? SelectedCostCenter = null;

        private VoucherLineItemViewModel newVoucherLine = new VoucherLineItemViewModel();
        private string validationMessage = string.Empty;

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

            Accounts = await _voucherService.GetTransactionalAccounts("");
            //cashAccounts = await _voucherService.GetTransactionalAccounts("WHERE SubAccOf='"+strCashInHandParent+ "'");
            //if (cashAccounts != null && cashAccounts.Any()) //Select Default 
            //{
            //    SelectedCashAccount = cashAccounts.FirstOrDefault(p => p.AccTitle == "Petty Cash") ?? cashAccounts.First();
            //    await SelectedResultChanged_CashAccounts(SelectedCashAccount);
            //}
            CostCenters = await _voucherService.GetValuesForDropDown("CostCenters", "EntryID", "CostCenter", " ORDER BY CostCenter");
            //TaxHeadTypes = await _voucherService.GetValuesForDropDown("EnumValues", "EnumValue", "EnumDescription", " WHERE EnumName='Voucher_Tax_Head_Type' ORDER BY EnumValue");
            
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

        [Parameter]
        public int p_bankId { get; set; }


        private async Task GetNextVoucherNo(DateTime selectedDate)
        {
            voucherViewModel.VchrNo = await VoucherService.GetNextVchrNo(selectedDate,"CRV");
        }

        private async Task OnDTChanged(ChangeEventArgs e)
        {
            await GetNextVoucherNo(voucherViewModel.DT);
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

            voucherViewModel.TaxRate = 15;
            voucherViewModel.DT = DateTime.Today;
            voucherViewModel.DueDate=DateTime.Today;
            //strCashBalance = string.Empty;
            voucherViewModel.Notes = string.Empty;
            voucherViewModel.Payee = string.Empty;
            voucherViewModel.LineItems.Clear();
            SelectedCostCenter = null;
            dDebitCreditDifference = 0;
            //SelectedCashAccount = null;
            /*voucherViewModel.ChequeDetails.Description=string.Empty;
            voucherViewModel.Notes = string.Empty;
            voucherViewModel.ChequeDetails.Payee=string.Empty;*/
            
            // Later i'll change amount & get from GD            
            await GetNextVoucherNo(voucherViewModel.DT);            

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
            if (EnableCostCenter && SelectedCostCenter == null)
            {
                NotificationServiceManager.ShowError("Please select Cost Center", "You need to select cost center from the drop-down.");
                return;
            }
            if (EnableCostCenter == false) 
            {
                
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
                CS_RefID = int.Parse(SelectedCostCenter?.DropDownValue_ID ?? "0"),
                CS_Description = SelectedCostCenter?.DropDownValue_Description ?? "",
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
            if (SelectedCostCenter != null)
                SelectedCostCenter = null;

            // Clear the validation message
            validationMessage = string.Empty;

        }

        private void RemoveVoucherLine(VoucherLineItemViewModel VLIVM)
        {
            voucherViewModel.LineItems.Remove(VLIVM);
            //voucherViewModel.ChequeDetails.Amount = voucherViewModel.TotalDebit - voucherViewModel.TotalCredit;
            //VoucherLineList.Remove(VLIVM);
        }

        
        // NOW DATE DATA INTO DATABSE 
        
        
        private async Task SaveVoucher()
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
                if (SelectedCashAccount == null)
                {
                    NotificationServiceManager.ShowWarning("Error", "Please select Cash Account.");
                    return;                    
                }
                                
                if (string.IsNullOrEmpty(voucherViewModel.Payee))
                {
                    NotificationServiceManager.ShowWarning("Error", "Please enter Received From.");
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
                dTotalCredit=voucherViewModel.TotalCredit+dDebitCreditDifference;
                if (dTotalDebit != dTotalCredit) 
                {
                    NotificationServiceManager.ShowWarning("Error", "Can't post transaction when Debit and Credit are not equal.");
                    return;
                }
                //We need to add Bank Account Transaction details as Credit to our voucher line.
                //systemLine = new VoucherLineItemViewModel
                //{
                //    AccNo = SelectedCashAccount.AccNo,
                //    AccTitle = SelectedCashAccount.AccTitle,
                //    Description = "Payment For "+voucherViewModel.Notes,  //voucherViewModel.ChequeDetails.Description,
                //    CS_RefID = 0,//int.Parse(SelectedCostCenter?.DropDownValue_ID ?? "0"),
                //    CS_Description = "",//SelectedCostCenter?.DropDownValue_Description ?? "",
                //    Debit = 0,
                //    Credit = dDebitCreditDifference,
                //    IsSystemLine=true
                //};
                //voucherViewModel.LineItems.Add(systemLine);

                /*voucherViewModel.ChequeDetails.Posted = true;
                voucherViewModel.ChequeDetails.Payment = true;
                voucherViewModel.ChequeDetails.Bounced = false;
                voucherViewModel.ChequeDetails.ChqIsDue = false;
                voucherViewModel.ChequeDetails.ClearanceDT = null;

                voucherViewModel.ChequeDetails.ChqPrintingDone = false;
                voucherViewModel.ChequeDetails.Vouchers_SNo = null;

                voucherViewModel.Payee = voucherViewModel.ChequeDetails.Payee;*/

                string strReturnVchrNo=await VoucherService.SaveVoucherAsync(voucherViewModel);

                isError = false;
                //successMessage = "Data Saved Successfully!";
                NotificationServiceManager.ShowSuccess("Saved", "Transaction saved successfully.");
                
                ReportNavigationService.PrintVoucher(strReturnVchrNo);
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

                ResetForm();
                
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
        private void ResetForm()
        {
            InitializeData();
            //Following 4 lines are used to clear fileinput
            bClearInputFile = true;
            StateHasChanged();
            bClearInputFile = false;
            StateHasChanged();

        }

        private void GoToIndexPage()
        {
            Navigation.NavigateTo("/Banks", true);
        }

        private IBrowserFile selectedFile;
        private void HandleFileSelected(InputFileChangeEventArgs e)
        {
            selectedFile = e.File;
        }

        private async Task UploadFileAsync(string primaryKey)
        {
            var client = ClientFactory.CreateClient("MyApiClient");
            if (selectedFile != null)
            {
                var uniqueFileName = $"{primaryKey}{Path.GetExtension(selectedFile.Name)}";
                uniqueFileName = uniqueFileName.Replace("/", "_");
                myFileExtension = Path.GetExtension(selectedFile.Name).ToLower();

                using (var content = new MultipartFormDataContent())
                {
                    var streamContent = new StreamContent(selectedFile.OpenReadStream(maxAllowedSize: 10_000_000)); // Set max allowed size
                    content.Add(streamContent, "file", uniqueFileName);
                    //var response = await Http.PostAsync("api/upload", content);
                    var response = await client.PostAsync("api/upload", content);

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new InvalidOperationException("File upload failed." + response.ToString());
                    }
                }
            }
        }
        
        
        private async Task<IEnumerable<ChartOfAccountsModel>> SearchCashAccounts(string searchText)
        {
            //return await Task.FromResult(submitTo.Where(x => x.AccTitle.ToLower().Contains(searchText.ToLower())).ToList());
            return await Task.FromResult(
                cashAccounts
                    .Where(x => x.AccTitle
                        .Contains(searchText, StringComparison.OrdinalIgnoreCase))
                    .ToList());
        }
        private async Task SelectedResultChanged_CashAccounts(ChartOfAccountsModel? selectedCashAccount)
        {
            SelectedCashAccount = selectedCashAccount;

            if (selectedCashAccount != null)
            {                
                decimal balance = await VoucherService.GetBalance(selectedCashAccount.AccNo, voucherViewModel.DT);
                strCashBalance = balance.ToString("N2");
                StateHasChanged();
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
            int iAccType = 0;
            iAccType = int.Parse(SelectedAccount.AccNo.Substring(0, 2));
            if (iAccType >= 31)
                EnableCostCenter = true;
            else
                EnableCostCenter = false;
            //newCIP.CardNo = SelectedCardNo.CardNo;
        }
        private async Task<IEnumerable<GenericDropDownModel>> SearchCostCenter(string searchText)
        {
            return await Task.FromResult(CostCenters.Where(x => x.DropDownValue_Description.ToLower().Contains(searchText.ToLower())).ToList());
        }
        private void SelectedResultChanged_CostCenter(GenericDropDownModel? selectedCostCenter)
        {
            SelectedCostCenter = selectedCostCenter;
        }
        
        private string GetAlertClass()
        {
            // Use the isError flag to return the appropriate Bootstrap class
            return isError ? "alert alert-danger" : "alert alert-success";
        }

        private string GetAlertTitle()
        {
            // Provide a different title based on the outcome
            return isError ? "Error!" : "Success!";
        }
        

        private async Task ToggleTaxModal()
        {
            if (showTaxModal == false) 
            {
                //before showing tax details check for following
                foreach (var line in voucherViewModel.LineItems)
                {
                    string strAccNo=line.AccNo;
                    if (strAccNo.Substring(0, 10) == "13-003-009") 
                    {
                        string strLast5 = strAccNo.Substring(strAccNo.Length - 5, 5);
                        if (strLast5 != "16005" && strLast5 != "16027") 
                        {
                            //fTaxRate = await _idbHelperService.getSingleValueasync<float>("Tax_Per", "Accounts", " WHERE AccNo='" + line.AccNo + "'");
                            fTaxRate = await _idbHelperService.GetSingleValueAsync<float>("Tax_Per", "Accounts", " WHERE AccNo='" + line.AccNo + "'");
                            fTaxAmount = (float) line.Debit;
                            break;
                        }
                    }
                    if (strAccNo== "13-003-009-16005")
                        fTaxAmount = (float)line.Debit;
                }
                showTaxModal = !showTaxModal;
            }
            else
                showTaxModal = false;

        }

        private void SaveTaxDetails()
        {
            // 1. Perform client-side validation if necessary.
            // (e.g., ensure TaxRate is > 0)

            // 2. The data is already bound to voucherViewModel, so no assignment is needed.

            // 3. Hide the modal
            showTaxModal = false;
        }
        private void HandleTaxTypeChange(ChangeEventArgs e)
        {
            // 1. Update TaxType
            if (int.TryParse(e.Value.ToString(), out int selectedTaxType))
            {
                voucherViewModel.TaxType = selectedTaxType;
                //Tax Type Option Button are 1 based index.
                if (voucherViewModel.TaxType == 1)  //Below Tax Limit..
                {
                    voucherViewModel.TaxRate = 0;
                    voucherViewModel.TaxAmount = 0;
                    voucherViewModel.HeadType = string.Empty;
                    voucherViewModel.Reason = string.Empty;
                }
                else if (voucherViewModel.TaxType == 2)     //Tax Deduction
                {
                    voucherViewModel.TaxRate = fTaxRate;
                    voucherViewModel.TaxAmount = (fTaxAmount* fTaxRate)/100;
                    voucherViewModel.Reason = string.Empty;
                }
                else if (voucherViewModel.TaxType == 3)  //Tax Exempted
                {
                    voucherViewModel.TaxRate = 0;
                    voucherViewModel.TaxAmount = 0;
                    voucherViewModel.HeadType = string.Empty;
                    voucherViewModel.Reason = string.Empty;
                }
                if (voucherViewModel.TaxType == 4) // 4 is for Tax Not Applicable
                {
                    voucherViewModel.TaxRate = 0;
                    voucherViewModel.TaxAmount = 0;
                    voucherViewModel.HeadType = string.Empty;
                    voucherViewModel.Reason = "Not Applicable";
                }
                else if (voucherViewModel.TaxType != 4 && voucherViewModel.Reason == "Not Applicable")
                {
                    // If they switch to another option, clear the "Not Applicable" text 
                    // so they can enter a real reason.
                    voucherViewModel.Reason = string.Empty;
                }
            }
        }

    }
}
