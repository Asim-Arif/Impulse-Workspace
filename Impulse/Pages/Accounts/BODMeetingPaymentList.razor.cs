using BlazorBootstrap;
using BlazorContextMenu;
using CIP.Pages;
using CIP.Pages.SelectionModals;
using CIP.Services;
using CIP.Services.Integrations;
using DataAccessLibrary;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using Humanizer;
using Microsoft.AspNetCore.Components;
using Microsoft.CodeAnalysis.Text;
using Microsoft.JSInterop;
using Radzen;

namespace CIP.Pages.Accounts
{
    public partial class BODMeetingPaymentList : ComponentBase
    {
        // Reference to the Modal Component
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

        private List<GenericDropDownModel> Banks = new List<GenericDropDownModel>();
        private GenericDropDownModel? SelectedBank = null;

        private List<BOD_Meeting_Payment_List_ViewModel> BODList { get; set; }=new List<BOD_Meeting_Payment_List_ViewModel>();
        private BOD_Meeting_Payment_List_ViewModel SelectedRow=new BOD_Meeting_Payment_List_ViewModel();
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
                Banks = await IVoucherService.GetValuesForDropDown("BankList", "BankID", "Bank+'{ ' + Branch +' '+ City + '}'");

                Banks.Insert(0, new GenericDropDownModel()
                {
                    DropDownValue_ID="0",
                    DropDownValue_Description="<All>"
                });

                if (Banks != null && Banks.Any()) //Select Default 
                {
                    SelectedBank = Banks.FirstOrDefault(p => p.DropDownValue_ID == "0") ?? Banks.First();
                    await SelectedResultChanged_Bank(SelectedBank);
                }
                DTFrom = DateTime.Today;
                DTTo = DateTime.Today;                

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

        
        private async Task<IEnumerable<GenericDropDownModel>> BanksSearch(string searchText)
        {
            return await Task.FromResult(
                Banks
                    .Where(x => x.DropDownValue_Description
                        .Contains(searchText,StringComparison.OrdinalIgnoreCase))
                    .ToList());
        }
        
        private async Task SelectedResultChanged_Bank(GenericDropDownModel? selectedBank)
        {
            SelectedBank = selectedBank;            
        }
        
        private async Task RefreshList()
        {
            IsLoading = true;
            try
            {                
                BODList.Clear();

                //var bodList = await AccountReportingAccess.GetLedgerData(CurrentAccount.AccNo,CurrentAccount.DTFrom,CurrentAccount.DTTo);
                var bodList = await AccountReportingAccessService.GetBODMeetingPaymentList(DTFrom,DTTo, SelectedStatus,int.Parse(SelectedBank.DropDownValue_ID));
                BODList = bodList.ToList();                
                
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
        private void OnRangeChanged(ChangeEventArgs e)
        {
            SelectedRangeValue = e.Value?.ToString();

            // Now logic based on the value
            switch (SelectedRangeValue)
            {
                case "0": // Today
                    DTFrom = DateTime.Today;
                    DTTo = DateTime.Today;
                    break;
                case "1": // Last 15 Days
                    DTFrom = DateTime.Today.AddDays(-15);
                    DTTo = DateTime.Today;
                    break;
                case "2": // Last 30 Days
                    DTFrom = DateTime.Today.AddDays(-30);
                    DTTo = DateTime.Today;
                    break;
                case "3": // Last 60 Days
                    DTFrom = DateTime.Today.AddDays(-60);
                    DTTo = DateTime.Today;
                    break;
                case "4": // Last 90 Days
                    DTFrom = DateTime.Today.AddDays(-90);
                    DTTo = DateTime.Today;
                    break;
                case "5": // Last 180 Days
                    DTFrom = DateTime.Today.AddDays(-180);
                    DTTo = DateTime.Today;
                    break;
                case "6": // Date Range
                          // We leave the dates as they are, letting the user edit them
                    break;
            }

            // Refresh your DirectorGridList based on these new dates
            //await FillDirectors(StartDate);
        }
        
        private async Task PostVoucherInERP(BOD_Meeting_Payment_List_ViewModel row) 
        {
            if (row == null) return;
            SelectedRow = row;

            if (!SelectedRow.PostedVchrNo.Equals("")) 
            {
                NotificationServiceManager.ShowWarning("Warning", "Already Posted.");
                return;
            }
            var chequeBooks = await IVoucherService.GetValuesForDropDown("VChqBooks", "ChqBookNo", "'Starting Cheque No. '+CAST(StartingFrom AS VARCHAR(50))+' {'+CAST(ChqsLeft AS VARCHAR(50))+' Cheques Left}' ", " WHERE AccNo='" + row.SIAL_Bank_Acc_No + "' AND ChqsLeft>0 AND Chq_Type=1");
            ChequeBookList = chequeBooks.Select(c => new GenericDropDownModel
            {
                DropDownValue_ID = c.DropDownValue_ID,
                DropDownValue_Description = c.DropDownValue_Description
            }).ToList();

            if (ChequeBookList.Any())
            {
                ChequeBookModal.Show();
            }
            else
            {
                NotificationServiceManager.ShowError("Error", $"\"No active cheque books found for this account.");                
            }
            /*var Data = e.Data as BOD_Meeting_Payment_List_ViewModel;
            SelectedRow = Data;
            NotificationServiceManager.ShowInformation("Info", $"Entry ID {SelectedRow.EntryID.ToString()}");*/
            //NotificationServiceManager.ShowInformation("Info", $"Selected Cheque Book No. {SelectedRow.EntryID.ToString()}");

        }
        protected async Task OnChequeBookSelected(string chequeBookNo)
        {
            // This runs when the user clicks 'OK' in the popup
            // Now you have both: selectedRow (the payment) and chequeBookNo (the book)
            Console.WriteLine($"Selected Cheque Book {chequeBookNo}");

            // TODO: Call your Bank API here

            try
            {
                VoucherViewModel voucherViewModel = new VoucherViewModel();
                voucherViewModel.UserName= AuditService.GetCurrentUserName();
                voucherViewModel.MachineName = AuditService.GetClientIpAddress();

                voucherViewModel.DT = DateTime.Today;
                voucherViewModel.VchrNo = await IVoucherService.GetNextVchrNo(DateTime.Today, "BPV");
                voucherViewModel.Notes = SelectedRow.Remarks;

                double SNo = await IVoucherService.GetNextSNo(voucherViewModel.DT, "Vouchers");                

                VoucherLineItemViewModel DebitLine= new VoucherLineItemViewModel();
                DebitLine.SNo = SNo;
                SNo += 1;
                DebitLine.VchrNo = voucherViewModel.VchrNo;
                DebitLine.VDate = voucherViewModel.DT;
                DebitLine.AccNo = SelectedRow.Debit_Acc_No;
                DebitLine.Description = SelectedRow.Remarks;
                DebitLine.Debit = SelectedRow.TotalDebit;
                DebitLine.Credit = 0;
                DebitLine.Balance = 0;

                voucherViewModel.LineItems.Add(DebitLine);

                
                if (SelectedRow.TotalTax > 0)
                {
                    VoucherLineItemViewModel TaxLine = new VoucherLineItemViewModel();                    
                    TaxLine.SNo= SNo;
                    SNo += 1;
                    TaxLine.VchrNo = voucherViewModel.VchrNo;
                    TaxLine.VDate = voucherViewModel.DT;
                    TaxLine.AccNo = SelectedRow.Tax_Acc_No;
                    TaxLine.Description = SelectedRow.Remarks;
                    TaxLine.Debit = 0;
                    TaxLine.Credit = SelectedRow.TotalTax;
                    TaxLine.Balance = 0;
                    voucherViewModel.LineItems.Add(TaxLine);
                }

                VoucherLineItemViewModel CreditLine = new VoucherLineItemViewModel();
                CreditLine.SNo = SNo;
                CreditLine.VchrNo = voucherViewModel.VchrNo;
                CreditLine.VDate = voucherViewModel.DT;
                CreditLine.AccNo = SelectedRow.SIAL_Bank_Acc_No;
                CreditLine.Description = SelectedRow.Remarks;
                CreditLine.Debit = 0;
                CreditLine.Credit = SelectedRow.TotalDebit- SelectedRow.TotalTax;
                CreditLine.Balance = 0;
                voucherViewModel.LineItems.Add(CreditLine);

                double dChqBookNo=double.Parse(chequeBookNo);
                IVoucherService.SaveBODMeetingPaymentVoucher(voucherViewModel, SelectedRow, dChqBookNo, DebitLine.SNo);

                NotificationServiceManager.ShowError("Success", "Voucher Posted Successfully.");
            }            
            
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Posting Failed", ex.Message);
            }
        }
        private async Task PostToBankPortal(BOD_Meeting_Payment_List_ViewModel row)
        {
            if (row == null) return;
            SelectedRow = row;

            if (SelectedRow.Status!=0)
            {
                NotificationServiceManager.ShowWarning("Warning", "Already Done.");
                return;
            }
            if (SelectedRow.PostedVchrNo.Equals(""))
            {
                NotificationServiceManager.ShowWarning("Warning", "Please Post in ERP before Posting in Portal.");
                return;
            }
            bool bReturValue= false;
            if (SelectedRow.BankID is 75 or 86) //Meezan Bank
            {
                bReturValue=await meezanBankService.PostToMeezanPortal(row.EntryID, row.Transaction_Type);
            }
            else if (SelectedRow.BankID is 90)  //Faysal Bank
            {
                //int iRecAffected=await faysalBankServices.ExportAndUploadExcelAsync(row.EntryID);
                int iRecAffected = await faysalBankServices.ExportAndUploadExcelForBOD(row.EntryID);
            }
            else 
            {
                NotificationServiceManager.ShowError("Error", "Bank API Not Found.");
                return;
            }

            if (bReturValue)
            {
                NotificationServiceManager.ShowSuccess("Success", "Done.");
            }
        }

    }
}