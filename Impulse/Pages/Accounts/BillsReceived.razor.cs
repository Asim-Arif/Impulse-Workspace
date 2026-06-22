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
    public partial class BillsReceived : ComponentBase
    {
        private bool ShowModal = false;
        private int EntryID_Edit = 0;
        private Bills_Received_ViewModel CurrentEntry = new();
        private int SerialFrom, SerialTo;

        private List<string> RecievedFromList = new();
        private List<string> RecievedByList = new();
        private List<string> HandedOverToList = new();

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
        private List<Bills_Received_ViewModel> ListToUse { get; set; }=new List<Bills_Received_ViewModel>();
        private Bills_Received_ViewModel SelectedRow =new Bills_Received_ViewModel();
        
        private DateTime DTFrom { get; set; } = DateTime.Now;
        private DateTime DTTo { get; set; } = DateTime.Now;

        private int SelectedStatus { get; set; } = 0;
        private string SelectedRangeValue { get; set; } = "0";
        protected override async Task OnInitializedAsync()
        {
            IsLoading = true;
            try
            {
                //Employees =await IVoucherService.GetValuesForDropDown("VEmp", "EmpID", "'{'+EmpID+'}' + Name + ' {'+DeptName+'}'", "WHERE Active=1");
                RecievedFromList = await IDBHelperService.GetListAsync<string>("DISTINCT ReceivedFrom", "BillsReceived", " ORDER BY ReceivedFrom");
                RecievedByList = await IDBHelperService.GetListAsync<string>("DISTINCT ReceivedBy", "BillsReceived", " ORDER BY ReceivedBy");
                HandedOverToList = await IDBHelperService.GetListAsync<string>("DISTINCT HandedOverTo", "BillsReceived", " ORDER BY HandedOverTo");

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
                
                var listtouse = await AccountReportingAccessService.GetBillsReceivedData(DTFrom,DTTo,SerialFrom,SerialTo);
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
            int iRefNo = await IDBHelperService.GetSingleValueAsync<int>("MAX(CAST(RefNo AS INT))", "BillsReceived", " WHERE ISNUMERIC(RefNo)=1");
            iRefNo += 1;
            

            CurrentEntry = new Bills_Received_ViewModel
            {
                DTIn = DateTime.Now,
                DTDue = DateTime.Today,
                DTHandedOver = DateTime.Now,
                RefNo=iRefNo
            };
            ShowModal = true;
        }

        private void OpenEdit(Bills_Received_ViewModel row) // Or your ContextMenu Data type
        {            
            CurrentEntry = new Bills_Received_ViewModel
            {
                EntryID = row.EntryID, 
                RefNo=row.RefNo,
                Party=row.Party,
                Description=row.Description,
                Amount=row.Amount,
                DTIn=row.DTIn,
                DTDue=row.DTDue,
                ReceivedFrom=row.ReceivedFrom,
                ReceivedBy=row.ReceivedBy,
                Objection=row.Objection,
                HandedOverTo=row.HandedOverTo,
                DTHandedOver=row.DTHandedOver,
                Status=row.Status,
                Remarks=row.Remarks                           
            };
            EntryID_Edit = row.EntryID;
            ShowModal = true;
        }
        private async Task DeleteEntry(Bills_Received_ViewModel row) 
        {
            int EntryID_To_Delete = row.EntryID;
            if (await Security.VerifyActionAsync("DeleteBillsReceived"))
            {
                string strSQL = "DELETE FROM BillsReceived WHERE EntryID=@EntryID";
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
                strSQL = @"INSERT INTO BillsReceived(RefNo,Party,Description,Amount,DTIn,DTDue,ReceivedFrom,ReceivedBy,Objection,HandedOverTo,DTHandedOver,Status,Remarks) VALUES(" +
                        "@RefNo,@Party,@Description,@Amount,@DTIn,@DTDue,@ReceivedFrom,@ReceivedBy,@Objection,@HandedOverTo,@DTHandedOver,@Status,@Remarks)";
            else
                strSQL = @"UPDATE BillsReceived SET RefNo=@RefNo,Party=@Party,Description=@Description,Amount=@Amount,DTIn=@DTIn,DTDue=@DTDue,ReceivedFrom=@ReceivedFrom"+
                        ",ReceivedBy=@ReceivedBy,Objection=@Objection,HandedOverTo=@HandedOverTo,DTHandedOver=@DTHandedOver,Status=@Status,Remarks=@Remarks" +
                        " WHERE EntryID=@EntryID_Edit";

            var parameters = new
            {
                RefNo= CurrentEntry.RefNo,
                Party= CurrentEntry.Party,
                Description= CurrentEntry.Description,
                Amount= CurrentEntry.Amount,
                DTIn= CurrentEntry.DTIn,
                DTDue= CurrentEntry.DTDue,
                ReceivedFrom= CurrentEntry.ReceivedFrom,
                ReceivedBy= CurrentEntry.ReceivedBy,
                Objection= CurrentEntry.Objection,
                HandedOverTo= CurrentEntry.HandedOverTo,
                DTHandedOver= CurrentEntry.DTHandedOver,
                Status= CurrentEntry.Status,
                Remarks= CurrentEntry.Remarks,                
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
            string strSelectionFormula = "";
            if (SerialFrom <= 0)
                strSelectionFormula = $"{{ImportantTasks.DTIn_DateOnly}}=#" + DTFrom.ToString("dd-MMM-yyyy") + "# to #" + DTTo.ToString("dd-MMM-yyyy") + "#";
            else
                strSelectionFormula = $"{{ImportantTasks.EntryID}}={SerialFrom} To {SerialTo}";

            var reportRequest = new ReportRequest
            {
                ReportName = ReportNames.BillsReceivedList,
                SelectionFormula = strSelectionFormula,
                FormulaValues = new Dictionary<string, object>
                {
                    { "DateRange", $"\"{strFromTo}\"" }
                }
            };

            ReportNavigationService.PrintReportAsync(reportRequest);                        
        }
        private void PrintEntry(Bills_Received_ViewModel row)
        {

            var reportRequest = new ReportRequest
            {
                ReportName = ReportNames.BillsReceived,
                SelectionFormula = $"{{ImportantTasks.EntryID}}={row.EntryID}"                
            };

            ReportNavigationService.PrintReportAsync(reportRequest);
        }
    }
}