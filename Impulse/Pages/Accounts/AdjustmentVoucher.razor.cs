using CIP.Services;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.JSInterop;
using Radzen;

namespace CIP.Pages.Accounts
{
    public partial class AdjustmentVoucher : ComponentBase, IDisposable
    {

        [Inject]
        private IVoucherService _voucherService { get; set; }
        [Inject]
        private IFinancialStatementService _financialstatementService { get; set; }
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

        private VouchersApprovalDetailViewModel currentLine;
        private bool IsModalVisible = false;
        private string strComputerName = string.Empty;
        private string strUserName = string.Empty;
        private string strCashInHandParent=string.Empty;
        private Vouchers_Main_Model voucherViewModel = new Vouchers_Main_Model();
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

        private Vouchers_Detail_Model newVoucherLine = new Vouchers_Detail_Model();
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
            strUserName=_auditService.GetCurrentUserName();

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
            InvokeAsync(() =>
            {
                successMessage = string.Empty;
                StateHasChanged();
            });
            _timer?.Dispose();
        }

        private async Task GetNextVoucherNo(DateTime selectedDate)
        {
            voucherViewModel.VchrNo = await _financialstatementService.GetNextVchrNo_Adjustment (selectedDate,"JV");
        }

        private async Task OnDTChanged(ChangeEventArgs e)
        {
            await GetNextVoucherNo(voucherViewModel.DT);
            StateHasChanged();
        }

        private async Task InitializeData()
        {
            isSaving = false;

            voucherViewModel.DT = DateTime.Today;
            
            //strCashBalance = string.Empty;
            //voucherViewModel.LineItems.Clear();
            voucherViewModel.Remarks = string.Empty;
            SelectedCostCenter = null;
            dDebitCreditDifference = 0;
            SelectedSubmitTo = null;
            
            // Later i'll change amount & get from GD            
            await GetNextVoucherNo(voucherViewModel.DT);
            StateHasChanged();

        }

        private async Task AddEntry()
        {

            if (SelectedAccount == null)
            {
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
            
            var lineToAdd = new Vouchers_Detail_Model
            {
                AccNo = SelectedAccount.AccNo,
                AccTitle = SelectedAccount.AccTitle,
                Description = newVoucherLine.Description,
                Debit = newVoucherLine.Debit,
                Credit = newVoucherLine.Credit
            };

            voucherViewModel.LineItems.Add(lineToAdd);
            
            dDebitCreditDifference = voucherViewModel.TotalDebit - voucherViewModel.TotalCredit;
            
            newVoucherLine.Description = string.Empty;
            newVoucherLine.Debit = 0;
            newVoucherLine.Credit = 0;

            SelectedAccount = null;
            validationMessage = string.Empty;

        }

        private void RemoveVoucherLine(Vouchers_Detail_Model VLIVM)
        {
            voucherViewModel.LineItems.Remove(VLIVM);
        }

        private async Task SaveVoucherMain()
        {
            if (isSaving)
                return;

            Vouchers_Detail_Model systemLine = null;
            try
            {
                isSaving = true;
                if (voucherViewModel.LineItems == null || !voucherViewModel.LineItems.Any())
                {
                    NotificationServiceManager.ShowWarning("Error", "No items added. Please add items before saving.");
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

                //Voucher_MainSaveResult voucher_MainSaveResult = await VoucherService.SaveVoucherMainAsync(voucherViewModel);
                var voucher_MainSaveResult = await _financialstatementService.SaveAdjustmentVoucher(voucherViewModel);

                isError = false;
                NotificationServiceManager.ShowSuccess("Saved", "Transaction saved successfully.");
                
                ResetForm();
               
                isSaving = false;
            }
            catch (Exception ex)
            {
                isError = true;
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

        private void ResetForm()
        {
            InitializeData();
            bClearInputFile = true;
            voucherViewModel.DT = DateTime.Today;
            voucherViewModel.Remarks = string.Empty;
            var voucher = new Vouchers_Main_Model
            {
                DT = DateTime.Now,
                LineItems = new List<Vouchers_Detail_Model> { new Vouchers_Detail_Model() }
            };
            StateHasChanged();
            if (newVoucherLine != null) 
            { 
                //newVoucherLine = null;
                voucherViewModel.LineItems = new List<Vouchers_Detail_Model>();
                //voucherViewModel.LineItems.Add(new Vouchers_Detail_Model());
            }
            StateHasChanged();
            bClearInputFile = false;
            StateHasChanged();
        }

        private void GoToIndexPage()
        {
            Navigation.NavigateTo("/Banks", true);
        }        
        private async Task<IEnumerable<GenericDropDownModel>> SearchSubmitTo(string searchText)
        {
            return await Task.FromResult(
                submitTo
                    .Where(x => x.DropDownValue_Description
                        .Contains(searchText, StringComparison.OrdinalIgnoreCase))
                    .ToList());
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
