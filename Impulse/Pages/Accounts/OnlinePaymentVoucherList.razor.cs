using BlazorBootstrap;
using BlazorContextMenu;
using Impulse.Services.Integrations;
using ClosedXML.Excel;
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
using Microsoft.IdentityModel.Tokens;
using Microsoft.JSInterop;
using Radzen;

namespace Impulse.Pages.Accounts
{
    public partial class OnlinePaymentVoucherList : ComponentBase
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
        

        private List<Online_Payment_Voucher_List_ViewModel> OVList { get; set; }=new List<Online_Payment_Voucher_List_ViewModel>();
        private Online_Payment_Voucher_List_ViewModel SelectedRow =new Online_Payment_Voucher_List_ViewModel();
        private DateTime DTFrom { get; set; } = DateTime.Now;
        private DateTime DTTo { get; set; } = DateTime.Now;

        private int SelectedStatus { get; set; } = 0;
        private int SelectedPaymentType { get; set; } = 0;
        
        private string SelectedRangeValue { get; set; } = "0";
        protected override async Task OnInitializedAsync()
        {
            IsLoading = true;
            try
            {
                IsLoading = true;
                Banks = await IVoucherService.GetValuesForDropDown("VBanks_For_OnlinePaymentVoucherList", "ID", "Description"," ORDER BY SortNo");

                /*AccountHeads.Insert(0, new GenericDropDownModel()
                {
                    DropDownValue_ID="0",
                    DropDownValue_Description="<All>"
                });*/

                if (Banks != null && Banks.Any()) //Select Default 
                {
                    SelectedBank = Banks.FirstOrDefault(p => p.DropDownValue_ID == "0") ?? Banks.First();
                    //await SelectedResultChanged_AccountHead(SelectedAccount);
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
            //if (string.IsNullOrWhiteSpace(searchText)) return AccountHeads;

            return await Task.FromResult(
                Banks
                    .Where(x => x.DropDownValue_Description
                        .Contains(searchText, StringComparison.OrdinalIgnoreCase))
                    .ToList());
            
        }

        private async Task SelectedResultChanged_Bank(GenericDropDownModel selectedBank)
        {
            SelectedBank = selectedBank;
        }
        
        private async Task RefreshList()
        {
            IsLoading = true;
            try
            {                
                OVList.Clear();

                //var bodList = await AccountReportingAccess.GetLedgerData(CurrentAccount.AccNo,CurrentAccount.DTFrom,CurrentAccount.DTTo);
                var bodList = await AccountReportingAccessService.GetOnlinePaymentVoucherList(DTFrom,DTTo, SelectedStatus,int.Parse(SelectedBank.DropDownValue_ID));
                OVList = bodList.ToList();                
                
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
        
        
        
        private async Task PostToBankPortal(Online_Payment_Voucher_List_ViewModel row)
        {
            if (row == null) return;
            SelectedRow = row;

            if (SelectedRow.Online_Posting)
            {
                NotificationServiceManager.ShowWarning("Warning", "Already Done.");
                return;
            }
            
            bool bReturValue= false;
            //This will be done later. 
            //HBL is not yet available.
            /*Faysal Bank
             * create ExportAndUploadExcelForOPV in Faysal Bank, create fetch in AccountsReportingDataAccess to fetch against vchr no
             * then use the existing function in Faysal Bank Services to handle and upate. use FaysalBankTransaction_Model Model.
            */
            /*Meezan Bank
             * Update PostToMeezanPortal Function in MeezanBankServices to accomoodate for Online Payment Voucher.
             * then call that function to post, 
            */
            if (SelectedRow.BankID is 75 or 86) //Meezan Bank
            {
                //bReturValue=await meezanBankService.PostToMeezanPortal(row.EntryID, row.Transaction_Type);
            }
            else if (SelectedRow.BankID is 90)  //Faysal Bank
            {
                //int iRecAffected=await faysalBankServices.ExportAndUploadExcelAsync(row.EntryID);
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
        private void ToggleAll(ChangeEventArgs e)
        {
            // 1. Get the state of the master checkbox from the event
            bool isChecked = (bool)(e.Value ?? false);

            // 2. Update every row in your grid
            foreach (var item in OVList)
            {
                item.IsChecked = isChecked;
            }

            // 3. (Optional) If your footer totals rely on these values, 
            // you might want to call StateHasChanged if the UI doesn't refresh automatically
            StateHasChanged();
        }
        private async Task PrintM1() 
        {
            var checkedItems = OVList.Where(x => x.IsChecked).ToList();
            if (!checkedItems.Any())
            {
                // Use your preferred notification service/JS Interop for MsgBox
                NotificationServiceManager.ShowError("Error", "Please select Entries.");
                return;
            }
            string firstM1No = checkedItems.First().M1No;
            if (checkedItems.Any(x => x.M1No != firstM1No))
            {
                NotificationServiceManager.ShowError("Error", "Can't Select Multiple M1.");                
                return;
            }
            int firstBankId = checkedItems.First().BankID;
            if (checkedItems.Any(x => x.BankID != firstBankId))
            {
                NotificationServiceManager.ShowError("Error", "Can't Proceed Different Banks Entries.");                
                return;
            }
            //string strEntryIDs = string.Join(",", checkedItems.Select(x => $"'{x.VchrNo}'"));
            var voucherNoList = checkedItems.Select(x => x.VchrNo).ToList();

            if (string.IsNullOrEmpty(firstM1No))
            {
                bool confirmed = await JS.InvokeAsync<bool>("confirm", "Are you sure to Generate New M1?");
                if (!confirmed) return;

                int month = DateTime.Now.Month;
                int year = DateTime.Now.Year;

                int nextM1No = await IDBHelperService.GetSingleValueAsync<int>("MAX(CAST(LEFT(M1No,LEN(M1No)-5) AS INT))", "VoucherInfo", " WHERE Month(M1DT)=" + month.ToString() + " AND YEAR(M1DT)=" + year.ToString());
                nextM1No += 1;

                string strM1No = $"{nextM1No}/{DateTime.Now:MMyy}";


                string sql = "UPDATE VoucherInfo SET M1No=@M1No, M1DT=@Date WHERE VchrNo IN @VoucherNos AND ISNULL(M1No,'')=''";
                IDBHelperService.ExecuteAsync(sql, new { M1No = strM1No, Date = DateTime.Today, VoucherNos = voucherNoList });
            }
            try
            {
                //IsLoading = true; // Start Animation

                string strSelectionFormula;
                /*strSelectionFormula = "{VCheque_M1.VDate}=#" + DTFrom.ToString("dd-MMM-yyyy") + "# TO #" + DTTo.ToString("dd-MMM-yyyy") + "#";                    
                strSelectionFormula += " AND {VCheque_M1.Online_Vchr}=TRUE AND {VCheque_M1.Online_Vchr_Type}=2";*/                    
                if (string.IsNullOrEmpty(firstM1No))
                    strSelectionFormula = "{VCheque_M1.VchrNo} IN [\"" + string.Join("\",\"", voucherNoList) + "\"]";
                else
                    strSelectionFormula = $"{{VCheque_M1.M1No}}='{firstM1No}'";

                bool bDuplicate=(string.IsNullOrEmpty(firstM1No)) ? false : true;
                string strDateRange = "'" + DTFrom.ToString("dd-MMM-yyyy") + " to " + DTTo.ToString("dd-MMM-yyyy") + "'";
                var reportRequest = new ReportRequest
                {
                    ReportName = ReportNames.Accounts.Cheques_M1,
                    SelectionFormula = strSelectionFormula,
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "Duplicate", bDuplicate },
                        { "DateRange", strDateRange },
                    }
                };
                await ReportNavigationService.PrintReportAsync(reportRequest);
                //await ReportNavigationService.PrintReportAsync(ReportNames.BOD_Meetings_Participants, strSelectionFormula);
            }
            finally
            {
                //IsLoading = false; // Stop Animation (placed in 'finally' so it stops even on error)
            }

            
        }
        private async Task DeleteM1(ItemClickEventArgs e) 
        {
            bool confirmed = await JS.InvokeAsync<bool>("confirm", "Are you sure to Delete M1?");
            if (!confirmed) return;

            var selectedRow = e.Data as Online_Payment_Voucher_List_ViewModel;
            if (selectedRow == null) return;

            string strM1No = selectedRow.M1No;
            int iRows=await IDBHelperService.ExecuteAsync("UPDATE VoucherInfo SET M1No=NULL,M1DT=NULL WHERE M1No=@M1No", new { M1No = strM1No});
            RefreshList();
        }

    }
}