using BlazorBootstrap;
using BlazorContextMenu;
using Impulse.Pages.SelectionModals;
using Impulse.Services.Integrations;
using DataAccessLibrary;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DocumentFormat.OpenXml.Spreadsheet;
using Humanizer;
using Impulse.Constants;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.CodeAnalysis.Text;
using Microsoft.JSInterop;
using Radzen;

namespace Impulse.Pages.Accounts
{
    public partial class LetterDispatchRegister : ComponentBase
    {
        private bool ShowModal = false;
        private int EntryID_Edit = 0;
        private Letter_Dispatch_Register_ViewModel CurrentEntry = new();
        
        private List<GenericDropDownModel> Employees = new();
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
        private SecurityService Security { get; set; }
        private List<Letter_Dispatch_Register_ViewModel> ListToUse { get; set; }=new List<Letter_Dispatch_Register_ViewModel>();
        private Letter_Dispatch_Register_ViewModel SelectedRow =new Letter_Dispatch_Register_ViewModel();
        
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

        private async Task RefreshList()
        {
            IsLoading = true;
            try
            {                
                ListToUse.Clear();
                
                var listtouse = await AccountReportingAccessService.GetLetterDispatchData(DTFrom,DTTo);
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

        private async Task<IEnumerable<GenericDropDownModel>> SearchEmployees(string searchText)
        {
            return await Task.FromResult(Employees
                .Where(x => x.DropDownValue_Description.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                            x.DropDownValue_ID.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                .ToList());
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
        }
        private async Task OpenAddNew()
        {
            EntryID_Edit = 0;
            int iRefNo = await IDBHelperService.GetSingleValueAsync<int>("MAX(CAST(RefNo AS INT))", "LetterDispatchRegister");
            CurrentEntry = new Letter_Dispatch_Register_ViewModel
            {
                DT = DateTime.Today,
                RefNo = iRefNo.ToString()
            };
            ShowModal = true;
        }

        private async Task OpenEdit(Letter_Dispatch_Register_ViewModel row) // Or your ContextMenu Data type
        {
            if (await Security.VerifyActionAsync("Edit_LDR"))
            {
                CurrentEntry = new Letter_Dispatch_Register_ViewModel
                {
                    EntryID = row.EntryID,
                    RefNo = row.RefNo,
                    DT = row.DT,
                    PartyName = row.PartyName,
                    Address = row.Address,
                    CellNo = row.CellNo,
                    CourierName = row.CourierName,
                    CourierNo = row.CourierNo
                };
                EntryID_Edit = row.EntryID;
                ShowModal = true;
            }
            
        }
        private async Task DeleteEntry(Letter_Dispatch_Register_ViewModel row) 
        {
            int EntryID_To_Delete = row.EntryID;
            if (await Security.VerifyActionAsync("Edit_LDR"))
            {
                string strSQL = "DELETE FROM LetterDispatchRegister WHERE EntryID=@EntryID";
                await IDBHelperService.ExecuteAsync(strSQL, new { EntryID = EntryID_To_Delete });
                ListToUse.Remove(row);
            }                
        }
        private void CloseModal()
        {
            ShowModal = false;
            EntryID_Edit = 0;
        } 

        private async Task SaveEntry()
        {
            // logic to call your Service.Save(CurrentEntry)
            string strSQL;
            if (EntryID_Edit == 0)
                strSQL = @"INSERT INTO LetterDispatchRegister(DT,RefNo,PartyName,Address,CellNo,CourierName,CourierNo) VALUES(" +
                        "@DT,@RefNo,@PartyName,@Address,@CellNo,@CourierName,@CourierNo)";
            else
                strSQL = @"UPDATE LetterDispatchRegister SET DT=@DT,RefNo=@RefNo,Partyname=@PartyName,Address=@Address,CellNo=@CellNo,CourierName=@CourierName,CourierNo=@CourierNo" +
                        " WHERE EntryID=@EntryID_Edit";

            var parameters = new
            {
                DT = CurrentEntry.DT,
                RefNo=CurrentEntry.RefNo,
                PartyName = CurrentEntry.PartyName,
                Address = CurrentEntry.Address,
                CellNo = CurrentEntry.CellNo,
                CourierName = CurrentEntry.CourierName,
                CourierNo = CurrentEntry.CourierNo,                
                EntryID_Edit = EntryID_Edit
            };
            await IDBHelperService.ExecuteAsync(strSQL,parameters);

            ShowModal = false;
            EntryID_Edit = 0;
            await RefreshList();
        }
        private void PrintReport() 
        {
            string strFromTo = DTFrom.ToString("dd-MMM-yyyy") + " to " + DTTo.ToString("dd-MMM-yyyy");            
            
            var reportRequest = new ReportRequest
            {
                ReportName = ReportNames.LetterDispatchRegister,
                SelectionFormula = $"{{VChequeReceivingReport.DT}}=#" +DTFrom.ToString("dd-MMM-yyyy")+"# to #"+DTTo.ToString("dd-MMM-yyyy")+"#",
                FormulaValues = new Dictionary<string, object>
                {
                    { "DateRange", $"\"{strFromTo}\"" }
                }
            };

            ReportNavigationService.PrintReportAsync(reportRequest);                        
        }
    }
}