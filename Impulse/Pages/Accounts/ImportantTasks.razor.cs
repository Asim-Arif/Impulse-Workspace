using BlazorBootstrap;
using BlazorContextMenu;
using CIP.Constants;
using CIP.Pages;
using CIP.Pages.SelectionModals;
using CIP.Services;
using CIP.Services.Integrations;
using DataAccessLibrary;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DocumentFormat.OpenXml.Spreadsheet;
using Humanizer;
using Microsoft.AspNetCore.Components;
using Microsoft.CodeAnalysis.Text;
using Microsoft.JSInterop;
using Radzen;

namespace CIP.Pages.Accounts
{
    public partial class ImportantTasks : ComponentBase
    {
        private bool ShowModal = false;
        private int EntryID_Edit = 0;
        private Important_Tasks_ViewModel CurrentEntry = new();

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
        private List<Important_Tasks_ViewModel> ListToUse { get; set; }=new List<Important_Tasks_ViewModel>();
        private Important_Tasks_ViewModel SelectedRow =new Important_Tasks_ViewModel();
        
        private DateTime DTFrom { get; set; } = DateTime.Now;
        private DateTime DTTo { get; set; } = DateTime.Now;

        private int SelectedStatus { get; set; } = 0;
        private string SelectedRangeValue { get; set; } = "0";
        protected override async Task OnInitializedAsync()
        {
            IsLoading = true;
            try
            {
                Employees =await IVoucherService.GetValuesForDropDown("VEmp", "EmpID", "'{'+EmpID+'}' + Name + ' {'+DeptName+'}'", "WHERE Active=1");
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
                
                var listtouse = await AccountReportingAccessService.GetImportantTasksData(DTFrom,DTTo);
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
        private void OpenAddNew()
        {
            EntryID_Edit = 0;
            CurrentEntry = new Important_Tasks_ViewModel
            {
                DT = DateTime.Today,
                DTDue = DateTime.Today
            };
            ShowModal = true;
        }

        private void OpenEdit(Important_Tasks_ViewModel row) // Or your ContextMenu Data type
        {            
            CurrentEntry = new Important_Tasks_ViewModel
            {
                EntryID = row.EntryID, 
                DT=row.DT,                
                Title = row.Title,
                Particulars=row.Particulars,
                DutyAssignedTo = row.DutyAssignedTo,
                Status=row.Status,
                DTDue=row.DTDue,
                OK = row.OK                
            };
            EntryID_Edit = row.EntryID;
            ShowModal = true;
        }
        private async Task DeleteEntry(Important_Tasks_ViewModel row) 
        {
            int EntryID_To_Delete = row.EntryID;
            if (await Security.VerifyActionAsync("DeleteImportantTask"))
            {
                string strSQL = "DELETE FROM ImportantTasks WHERE EntryID=@EntryID";
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
                strSQL = @"INSERT INTO ImportantTasks(DT,Title,Particulars,DutyAssignedTo,Status,DTDue,OK) VALUES(" +
                        "@DT,@Title,@Particulars,@DutyAssignedTo,@Status,@DTDue,@OK)";
            else
                strSQL = @"UPDATE ImportantTasks SET DT=@DT,Title=@Title,Particulars=@Particulars,DutyAssignedTo=@DutyAssignedTo,Status=@Status,DTDue=@DTDue" +
                        ",OK=@OK WHERE EntryID=@EntryID_Edit";

            var parameters = new
            {
                DT = CurrentEntry.DT,
                Title = CurrentEntry.Title,
                Particulars = CurrentEntry.Particulars,
                DutyAssignedTo = CurrentEntry.DutyAssignedTo,
                Status = CurrentEntry.Status,
                DTDue = CurrentEntry.DTDue,
                OK = CurrentEntry.OK,
                EntryID_Edit = EntryID_Edit
            };
            await IDBHelperService.ExecuteAsync(strSQL,parameters);

            ShowModal = false;
            EntryID_Edit = 0;
            await RefreshList();
        }
        private void PrintList() 
        {
            string strFromTo = DTFrom.ToString("dd-MMM-yyyy") + " to " + DTTo.ToString("dd-MMM-yyyy");            
            
            var reportRequest = new ReportRequest
            {
                ReportName = ReportNames.ImportantTasksList,
                SelectionFormula = $"{{ImportantTasks.DT}}=#" +DTFrom.ToString("dd-MMM-yyyy")+"# to #"+DTTo.ToString("dd-MMM-yyyy")+"#",
                FormulaValues = new Dictionary<string, object>
                {
                    { "DateRange", $"\"{strFromTo}\"" }
                }
            };

            ReportNavigationService.PrintReportAsync(reportRequest);                        
        }
        private void PrintEntry(Important_Tasks_ViewModel row)
        {

            var reportRequest = new ReportRequest
            {
                ReportName = ReportNames.ImportantTasks,
                SelectionFormula = $"{{ImportantTasks.EntryID}}={row.EntryID}"                
            };

            ReportNavigationService.PrintReportAsync(reportRequest);
        }
    }
}