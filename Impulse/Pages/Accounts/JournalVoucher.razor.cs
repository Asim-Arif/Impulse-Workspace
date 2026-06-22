using CIP.Services;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.JSInterop;
using Radzen;

namespace CIP.Pages.Accounts
{
    public partial class JournalVoucher : ComponentBase, IDisposable
    {
        private List<VAD_Voucher_DetailViewModel> pendingJVs = new();
        private VouchersApprovalDetailViewModel currentLine;
        private bool IsModalVisible = false;
        private string strComputerName = string.Empty;
        private string strUserName = string.Empty;
        private string strCashInHandParent=string.Empty;
        private Vouchers_MainViewModel voucherViewModel = new Vouchers_MainViewModel();
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
        
        private List<GenericDropDownModel> submitTo = new List<GenericDropDownModel>();
        private List<ChartOfAccountsModel> Accounts = new List<ChartOfAccountsModel>();
        private List<GenericDropDownModel> CostCenters = new List<GenericDropDownModel>();
        private List<GenericDropDownModel> TaxHeadTypes = new List<GenericDropDownModel>();
        
        private GenericDropDownModel? SelectedSubmitTo = null;
        private ChartOfAccountsModel? SelectedAccount = null;        
        private GenericDropDownModel? SelectedCostCenter = null;

        private VouchersApprovalDetailViewModel newVoucherLine = new VouchersApprovalDetailViewModel();
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

            strCashInHandParent = await _idbHelperService.getSingleStringValue("DataValue", "GeneralData", "WHERE DataName='CashInHandParent'");

            Accounts = await _voucherService.GetTransactionalAccounts("WHERE SubAccOf<>'"+strCashInHandParent+"'");
            string strHODCondition = " WHERE (UserName<>'" + strUserName + "' AND InActive=0 AND UserRole NOT IN(''))";
            submitTo = await _voucherService.GetValuesForDropDown("Users INNER JOIN VEmps ON Users.EmpID=VEmps.EmpID", "UserName", "VEmps.Name+' ('+VEmps.Designation+')'", strHODCondition+" ORDER BY VEmps.PayGroup_EntryID DESC");
            
            CostCenters = await _voucherService.GetValuesForDropDown("CostCenters", "EntryID", "CostCenter", " ORDER BY CostCenter");
            TaxHeadTypes = await _voucherService.GetValuesForDropDown("EnumValues", "EnumValue", "EnumDescription", " WHERE EnumName='Voucher_Tax_Head_Type' ORDER BY EnumValue");
            
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
            voucherViewModel.VchrNo = await VoucherService.GetNextVchrNo_Approval(selectedDate,"JV");
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

            voucherViewModel.TaxRate = 15;
            voucherViewModel.DT = DateTime.Today;
            
            //strCashBalance = string.Empty;
            voucherViewModel.LineItems.Clear();
            voucherViewModel.ReceivedFrom = string.Empty;
            voucherViewModel.SlipNo = string.Empty;
            voucherViewModel.Remarks = string.Empty;
            SelectedCostCenter = null;
            dDebitCreditDifference = 0;
            SelectedSubmitTo = null;
            /*voucherViewModel.ChequeDetails.Description=string.Empty;
            voucherViewModel.Notes = string.Empty;
            voucherViewModel.ChequeDetails.Payee=string.Empty;*/
            
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
            
            var lineToAdd = new VouchersApprovalDetailViewModel
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

        private void RemoveVoucherLine(VouchersApprovalDetailViewModel VLIVM)
        {
            voucherViewModel.LineItems.Remove(VLIVM);
            //voucherViewModel.ChequeDetails.Amount = voucherViewModel.TotalDebit - voucherViewModel.TotalCredit;
            //VoucherLineList.Remove(VLIVM);
        }
        private void AddDetails(VouchersApprovalDetailViewModel VLIVM)
        {
            OpenAdjustmentModal(VLIVM);
        }
        private void SaveAdjustments() 
        {
            var lineValue = (decimal)Math.Abs((currentLine?.Debit ?? 0) - (currentLine?.Credit ?? 0));
            var totalAdjusting = (decimal)pendingJVs.Sum(x => x.Amount ?? 0);
            if (totalAdjusting > lineValue) 
            {
                NotificationServiceManager.ShowError("Invalid Total", "Can't enter more than Voucher Amount.");
                return; // Stop the save
            }
            foreach (var jv in pendingJVs.Where(x => x.Amount > 0))
            {
                if (jv.Amount > jv.BalanceAmount)
                {
                    NotificationServiceManager.ShowError("Invalid", "Can't enter more than available balance");
                    return; // Stop the save
                }
            }

            currentLine.AdjustmentDetails = pendingJVs
                .Where(x => x.Amount > 0)
                .ToList();

                    // Optional: Update the main line description or a badge to show it's "Adjusted"
            if (currentLine.AdjustmentDetails.Any())
            {
                // For example, you could update the description to show it's linked
                // currentLine.Description += " (Adjusted)";
            }

            IsModalVisible = false;
            StateHasChanged();
        }
        // NOW DATE DATA INTO DATABSE 


        private async Task SaveVoucherMain()
        {
            if (isSaving)
                return;

            VouchersApprovalDetailViewModel systemLine = null;
            try
            {
                isSaving = true;
                // Ensure cipList has items
                if (voucherViewModel.LineItems == null || !voucherViewModel.LineItems.Any())
                {
                    NotificationServiceManager.ShowWarning("Error", "No items added. Please add items before saving.");
                    return;
                }                
                if (SelectedSubmitTo == null)
                {
                    NotificationServiceManager.ShowWarning("Error", "Please select Submit To.");
                    return;                    
                }
                                
                
                if (string.IsNullOrEmpty(voucherViewModel.Remarks))
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
                
                voucherViewModel.Attachment_Type = ".pdf";

                Voucher_MainSaveResult voucher_MainSaveResult = await VoucherService.SaveVoucherMainAsync(voucherViewModel);

                if (!bClearInputFile)
                {
                    await UploadFileAsync(voucher_MainSaveResult.EntryID);
                    {
                        voucherViewModel.Attachment_Type = myFileExtension;
                    }
                }

                isError = false;
                //successMessage = "Data Saved Successfully!";
                NotificationServiceManager.ShowSuccess("Saved", "Transaction saved successfully.");
                
                //ReportNavigationService.PrintVoucher(strReturnVchrNo);
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

        private async Task UploadFileAsync(int iVM_EntryID)
        {
            var client = ClientFactory.CreateClient("MyApiClient");
            if (selectedFile != null)
            {
                string extension = Path.GetExtension(selectedFile.Name);
                string uniqueFileName = $"BRV_{iVM_EntryID}{extension}";
                
                uniqueFileName = uniqueFileName.Replace("/", "_");                

                using (var content = new MultipartFormDataContent())
                {
                    var streamContent = new StreamContent(selectedFile.OpenReadStream(maxAllowedSize: 10_000_000)); // Set max allowed size
                    content.Add(streamContent, "file", uniqueFileName);                    
                    var response = await client.PostAsync("api/upload", content);

                    if (!response.IsSuccessStatusCode)
                    {
                        var errorDetails = await response.Content.ReadAsStringAsync();
                        throw new InvalidOperationException("File upload failed." + response.ToString());
                    }
                }
            }
        }
        
        
        private async Task<IEnumerable<GenericDropDownModel>> SearchSubmitTo(string searchText)
        {
            //return await Task.FromResult(submitTo.Where(x => x.AccTitle.ToLower().Contains(searchText.ToLower())).ToList());
            return await Task.FromResult(
                submitTo
                    .Where(x => x.DropDownValue_Description
                        .Contains(searchText, StringComparison.OrdinalIgnoreCase))
                    .ToList());
        }
        private async Task SelectedResultChanged_SubmitTo(GenericDropDownModel? selectedSubmitTo)
        {
            SelectedSubmitTo = selectedSubmitTo;
            if (SelectedSubmitTo != null && SelectedSubmitTo.DropDownValue_ID != null) 
            {
                voucherViewModel.ForwardTo = SelectedSubmitTo.DropDownValue_ID.ToString();
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

        private async Task OpenAdjustmentModal(VouchersApprovalDetailViewModel selectedLine)
        {
            currentLine = selectedLine;

            // Equivalent to RefreshVouchersList logic:
            var data = await VoucherService.GetPendingJVsByAccount(selectedLine.AccNo);

            // Map the fetched data to our UI list
            pendingJVs = data.Select(x => new VAD_Voucher_DetailViewModel
            {
                Sno = x.Sno,
                VchrNo = x.VchrNo,
                VchrAmount= x.VchrAmount,
                Amount = 0,
                BalanceAmount = x.BalanceAmount,
                AdjAmount = x.AdjAmount, // This is ITM.SubItems(4) = 0 in your VB code (User Input)
                AccNo = x.AccNo,
                RefID = currentLine.EntryID // Link back to the line item
            }).ToList();

            if (currentLine.AdjustmentDetails != null && currentLine.AdjustmentDetails.Any())
            {
                foreach (var item in pendingJVs)
                {
                    // Look for a matching Voucher Number in our saved list
                    var existing = currentLine.AdjustmentDetails
                        .FirstOrDefault(d => d.VchrNo == item.VchrNo && d.Sno == item.Sno);

                    if (existing != null)
                    {
                        // Restore the value the user entered last time
                        item.Amount = existing.Amount;
                    }
                }
            }
            // Show your modal here
            IsModalVisible = true;
            StateHasChanged();
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
