using CIP.Constants;
using CIP.Services;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.JSInterop;
using Radzen;

namespace CIP.Pages.Accounts
{
    public partial class BODMeetingPayment : ComponentBase
    {
        private bool IsLoading = false;
        private bool IsDataLoading = false;
        private bool showBulkModal = false;
        private int globalDebit;
        private int globalTax;

        private string strComputerName=string.Empty;
        private string strUserName=string.Empty;
        private Online_PVModelView voucherViewModel = new Online_PVModelView();
        private SIALDirectorViewModel sIALDirectorViewModel = new SIALDirectorViewModel();
        private float fTaxRate=0;
        private float fTaxAmount = 0;
        private string VoucherNo { get; set; }=string.Empty;
        private bool bClearInputFile = false;        
        private bool EnableCostCenter = false;
        private string successMessage = string.Empty;
        private bool isError = false;
        private bool showTaxModal = false;
        private string myFileExtension = string.Empty;
        private bool isSaving = false;
        private DateTime selectedDate = DateTime.Today;
        private string strBankBalance { get; set; } = "";
        private List<SIALDirectorPaymentViewModel> DirectorGridList = new List<SIALDirectorPaymentViewModel>();
        private List<SIALDirectorViewModel> DirectorsList=new List<SIALDirectorViewModel>();
        
        private List<BankAccountInfo> bankAccounts = new List<BankAccountInfo>();
        private List<ChartOfAccountsModel> debitAccounts = new List<ChartOfAccountsModel>();
        private List<ChartOfAccountsModel> taxAccounts = new List<ChartOfAccountsModel>();

        
        private BankAccountInfo? SelectedBankAccount = null;
        private ChartOfAccountsModel? SelectedDebitAccount = null;
        private ChartOfAccountsModel? SelectedTaxAccount = null;
        private SIALDirectorViewModel? SelectedDirector = null;

        private VoucherLineItemViewModel newVoucherLine = new VoucherLineItemViewModel();
        private string validationMessage = string.Empty;

        
        protected override async Task OnInitializedAsync()
        {
            await InitializeData();

            strComputerName= _auditService.GetClientIpAddress();
            strUserName=  _auditService.GetCurrentUserName();

            bankAccounts = await _voucherService.GetBankAccounts();
            debitAccounts = await _voucherService.GetTransactionalAccounts();
            if (debitAccounts != null && debitAccounts.Any()) //Select Default 
            {
                SelectedDebitAccount = debitAccounts.FirstOrDefault(p => p.AccNo == "41-017-13001") ?? debitAccounts.First();
                await SelectedResultChanged_DebitAccount(SelectedDebitAccount);
            }
            taxAccounts = await _voucherService.GetTransactionalAccounts();
            if (taxAccounts != null && taxAccounts.Any()) //Select Default 
            {
                SelectedTaxAccount = taxAccounts.FirstOrDefault(p => p.AccNo == "13-003-009-16011") ?? taxAccounts.First();
                await SelectedResultChanged_TaxAccount(SelectedTaxAccount);
            }
        }
        
        [Inject]
        protected IMeezanBankService _meezanBankService { get; set; } = default!;
        [Inject]
        private IVoucherService _voucherService { get; set; } = null!;
        [Inject]
        private ICommonServices _commonServices { get; set; }
        [Inject]
        protected ILogger<BODMeetingPayment> _logger { get; set; } = default!;
        [Inject]
        private IAuditService _auditService { get; set; } = null!;

        [Inject]
        private IDBHelperService _idbHelperService { get; set; } = null!;

        [Inject]
        private INotificationService NotificationServiceManager { get; set; } = null!;

        [Inject]
        private IReportNavigationService ReportNavigationService { get; set; } = null!;
        
        [Inject]
        protected IJSRuntime JSRuntime { get; set; } = default!;
        [Parameter]
        public int p_bankId { get; set; }


        private async Task FillDirectors(DateTime selectedDate)
        {
            try 
            {
                IsDataLoading = true;
                DirectorsList = await _voucherService.GetDirectorsByBODMeetingDate(selectedDate);

                DirectorGridList = DirectorsList.Select(d => new SIALDirectorPaymentViewModel
                {
                    Director = d,
                    IsChecked = false,
                    IsProcessed = d.SD_RefID != null,
                    Debit = 0,
                    Tax = 0
                }).ToList();
                //await Task.Delay(100); // Small delay to ensure the DOM refreshes
            }
            finally
            {
                IsDataLoading = false;
            }
        }

        private async Task OnDTChanged(ChangeEventArgs e)
        {            
            await FillDirectors(voucherViewModel.DT);
        }

        private async Task InitializeData()
        {
            isSaving = false;

            DateTime? LastMeetingDate= await _idbHelperService.getSingleDateValue("MAX(DT)", "BOD_Meetings");
            voucherViewModel.DT = LastMeetingDate ?? DateTime.Today;
            
            strBankBalance = string.Empty;
            voucherViewModel.LineItems.Clear();            
            SelectedBankAccount = null;            
            voucherViewModel.Remarks = string.Empty;                        
            // Later i'll change amount & get from GD 
            await FillDirectors(voucherViewModel.DT);
        }

        private async Task SaveVoucher()
        {
            if (isSaving)
                return;

            try
            {
                isSaving = true;
                var selectedDirectors = DirectorGridList.Where(x => x.IsChecked).ToList();
                if (!selectedDirectors.Any())
                {
                    NotificationServiceManager.ShowWarning("Error", "Please select Director from the table.");
                    return;
                }
                if (SelectedBankAccount == null)
                {
                    NotificationServiceManager.ShowWarning("Error", "Please select Bank Account.");
                    return;                    
                }
                if (SelectedDebitAccount == null)
                {
                    NotificationServiceManager.ShowWarning("Error", "Please select Debit Account.");
                    return;
                }
                if (SelectedDebitAccount == null)
                {
                    NotificationServiceManager.ShowWarning("Error", "Please select Tax Account.");
                    return;
                }

                if (string.IsNullOrEmpty(voucherViewModel.Remarks))
                {
                    NotificationServiceManager.ShowWarning("Error", "Please enter Remarks.");
                    return;
                }
                if (voucherViewModel.Transaction_Type==null)
                {
                    NotificationServiceManager.ShowWarning("Error", "Please select Payment Mode.");
                    return;
                }

                isError = false;
                //successMessage = "Data Saved Successfully!";
                int iBOD_Meetings_RefID = _idbHelperService.getSingleIntValue("EntryID", "BOD_Meetings","WHERE DT=@DT",new {DT=voucherViewModel.DT });
                voucherViewModel.BOD_Meetings_RefID = iBOD_Meetings_RefID;
                voucherViewModel.UserName = strUserName;
                voucherViewModel.MachineName = strComputerName;
                voucherViewModel.SIAL_Bank_Acc_No = SelectedBankAccount.AccNo;
                voucherViewModel.Debit_Acc_No = SelectedDebitAccount.AccNo;
                voucherViewModel.Tax_Acc_No = SelectedTaxAccount.AccNo;

                var details = selectedDirectors.Select(d => new Online_PV_DetailsModelView
                {                    
                    SD_RefID=d.Director.EntryID,
                    Bank_Acc_No=d.Director.Bank_Account_No,
                    Bank_Code=d.Director.Bank_Code,
                    Bank_Acc_Title=d.Director.Bank_Account_Title,
                    Bank_Name=d.Director.Bank_Name,
                    Debit = d.Debit,
                    Tax = d.Tax,
                    
                }).ToList();

                int iPV_EntryID=await _voucherService.SaveOnlinePaymentVoucher(voucherViewModel, details);

                NotificationServiceManager.ShowSuccess("Saved", "Transaction saved successfully.");

                //ReportNavigationService.PrintVoucher(strReturnVchrNo);
                
                ResetForm();
                
                //VoucherLineList.Clear();
                isSaving = false;
            }
            catch (Exception ex)
            {
                isError = true;                
                NotificationServiceManager.ShowError("Error", $"{ex.Message}");
                
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

        private async Task<IEnumerable<BankAccountInfo>> SearchBankAccounts(string searchText)
        {
            return await Task.FromResult(
                bankAccounts
                    .Where(x => x.AccTitle
                        .Contains(searchText,StringComparison.OrdinalIgnoreCase))
                    .ToList());
        }
        private async Task<IEnumerable<ChartOfAccountsModel>> SearchDebitAccount(string searchText)
        {
            return await Task.FromResult(
                debitAccounts
                    .Where(x => x.AccTitle
                        .Contains(searchText, StringComparison.OrdinalIgnoreCase))
                    .ToList());
        }
        private async Task<IEnumerable<ChartOfAccountsModel>> SearchTaxAccount(string searchText)
        {
            return await Task.FromResult(
                taxAccounts.Where(x => x.AccTitle
                    .Contains(searchText, StringComparison.OrdinalIgnoreCase))
                    .ToList());
        }
        private async Task<IEnumerable<SIALDirectorViewModel>> SearchDirectors(string searchText)
        {
            return await Task.FromResult(
                DirectorsList.Where(x => x.DirectorName
                    .Contains(searchText, StringComparison.OrdinalIgnoreCase))
                    .ToList());
        }
        private void SelectedResultChanged_BankAccount(BankAccountInfo? selectedBankAccount)
        {
            SelectedBankAccount = selectedBankAccount;
        }
        private async Task SelectedResultChanged_DebitAccount(ChartOfAccountsModel? selectedDebitAccount)
        {
            SelectedDebitAccount = selectedDebitAccount;
        }
        private async Task SelectedResultChanged_TaxAccount(ChartOfAccountsModel? selectedTaxAccount)
        {
            SelectedTaxAccount = selectedTaxAccount;
        }
        private async Task SelectedResultChanged_Director(SIALDirectorViewModel? selectedDirector)
        {
            SelectedDirector = selectedDirector;
            var target = DirectorGridList.FirstOrDefault(x => x.Director.EntryID == SelectedDirector.EntryID);
            if (target != null)
            {
                // 2. Auto-check the box
                //target.IsChecked = true;

                // 3. Highlight the row
                target.IsHighlighted = true;

                await JSRuntime.InvokeVoidAsync("eval", $"document.getElementById('row-{target.Director.EntryID}')?.scrollIntoView({{behavior: 'smooth', block: 'center'}})");
                // Or use: await JSRuntime.InvokeVoidAsync("eval", $"document.getElementById('row-{target.Director.FolioNo}').scrollIntoView({{behavior: 'smooth', block: 'center'}})");
                // 4. Optional: Clear the highlight after 3 seconds so it's not permanent
                await Task.Delay(3000).ContinueWith(_ => {
                    target.IsHighlighted = false;
                    InvokeAsync(StateHasChanged);
                });
            }

        }
        private void HandlePayMethodChange(ChangeEventArgs e)
        {
            // 1. Update TaxType
            if (int.TryParse(e.Value.ToString(), out int selectedTaxType))
            {
                voucherViewModel.Transaction_Type = selectedTaxType;
                //Tax Type Option Button are 1 based index.
                if (voucherViewModel.Transaction_Type == 0)  //Online Funds Transfer
                {
                    foreach (var item in DirectorGridList)
                    {
                        item.IsChecked = false;
                        if (!string.IsNullOrWhiteSpace(item.Director.Bank_Account_Title) && item.IsProcessed==false)
                        {
                            item.IsChecked = true;
                        }
                    }
                }
                else if (voucherViewModel.Transaction_Type == 1)     //Corporate Cheque
                {
                    foreach (var item in DirectorGridList)
                    {
                        item.IsChecked = false;
                        if (string.IsNullOrWhiteSpace(item.Director.Bank_Account_Title) && item.IsProcessed == false)
                        {
                            item.IsChecked = true;
                        }
                    }
                }
                
                
            }
        }
        private void ToggleAll(ChangeEventArgs e)
        {
            // 1. Get the state of the master checkbox from the event
            bool isChecked = (bool)(e.Value ?? false);

            // 2. Update every row in your grid
            foreach (var item in DirectorGridList)
            {
                item.IsChecked = isChecked;
            }

            // 3. (Optional) If your footer totals rely on these values, 
            // you might want to call StateHasChanged if the UI doesn't refresh automatically
            StateHasChanged();
        }
        private void OpenBulkModal()
        {
            // Prevent opening if no one is selected
            /*if (!DirectorGridList.Any(x => x.IsChecked))
            {
                // Optional: show a notification that no directors are selected
                return;
            }*/
            showBulkModal = true;
            //StateHasChanged();
        }

        private void CloseModal()
        {
            showBulkModal = false;
        }

        private void ApplyBulkValues()
        {
            foreach (var item in DirectorGridList)
            {
                item.Debit = globalDebit;
                item.Tax = globalTax;
            }

            showBulkModal = false; // Close after applying
                                   // StateHasChanged(); // Only needed if UI doesn't refresh
        }
        private async Task HandleExcelUpload(InputFileChangeEventArgs e)
        {
            var file = e.File;
            if (file != null)
            {
                try
                {
                    await _commonServices.ProcessBODMeetingAttendanceExcel(file, voucherViewModel.DT);
                    // Refresh the grid after import
                    await FillDirectors(voucherViewModel.DT);
                    // Show Success Message
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error importing Excel file for date {Date}", voucherViewModel.DT);
                }
            }
        }
        private async Task FetchTitle()
        {
            string strDebit_Acc_No = "0101783774";
            foreach (var item in DirectorGridList)
            {
                string strTransaction_Type;
                if (item.Director.Bank_Name?.Contains("Meezan", StringComparison.OrdinalIgnoreCase) ?? false)
                    strTransaction_Type = "InternalFundsTransfer_TitleFetch";
                else
                    strTransaction_Type = "InterBankFundTransfer_TitleFetch";

                string strResponse = await _meezanBankService.FetchTitleMeezan(strDebit_Acc_No, item.Director.Bank_Account_No ?? "", strTransaction_Type, item.Director.Bank_Code ?? "");
                
                string responseCode = _commonServices.GetValueFromXMLString(strResponse, "ResponseCode");
                string responseDescription = _commonServices.GetValueFromXMLString(strResponse, "ResponseDescription");
                string responseTxnNo = _commonServices.GetValueFromXMLString(strResponse, "UnqiueTxnNo");

                string accountTitle = responseDescription.Replace("0;","");
                item.Director.Bank_Account_Title = accountTitle;
                item.IsChecked = true;
                item.IsOnlineTitleFetched = true;
            }
        }
        private async Task PrintAttendees(bool bProcessed) 
        {
            try
            {
                IsLoading = true; // Start Animation

                int iBOD_Meetings_RefID = _idbHelperService.getSingleIntValue("EntryID", "BOD_Meetings", "WHERE DT=@DT", new { DT = voucherViewModel.DT });
                string strSelectionFormula = "{VBOD_Meetings_Participants.BOD_Meetings_RefID}=" + iBOD_Meetings_RefID.ToString();
                if (bProcessed)
                    strSelectionFormula += " AND {VBOD_Meetings_Participants.SD_RefID_O_PV_D}>0";

                var reportRequest = new ReportRequest
                {
                    ReportName = ReportNames.BOD_Meetings_Participants,
                    SelectionFormula = strSelectionFormula,
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "Test", "\"Value from Front end\"" }
                    }
                };
                await ReportNavigationService.PrintReportAsync(reportRequest);
                //await ReportNavigationService.PrintReportAsync(ReportNames.BOD_Meetings_Participants, strSelectionFormula);
            }
            finally 
            {
                IsLoading = false; // Stop Animation (placed in 'finally' so it stops even on error)
            }
            
        }

    }
}
