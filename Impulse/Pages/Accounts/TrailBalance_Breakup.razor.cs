using BlazorBootstrap;
using BlazorContextMenu;
using DataAccessLibrary;
using DataAccessLibrary.DAC.Accounts;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DocumentFormat.OpenXml.Drawing;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Microsoft.IdentityModel.Tokens;
using Microsoft.JSInterop;
using Microsoft.VisualBasic;
using Radzen;
using System.Collections.ObjectModel;
using System.Linq;
using System.Security.AccessControl;
using System.Threading.Tasks;
using OpenXmlMenu = DocumentFormat.OpenXml.Office2010.CustomUI;


namespace Impulse.Pages.Accounts
{
    public partial class TrailBalance_Breakup : ComponentBase, IDisposable
    {
        private TrailBalance_BreakupModel statementmodel = new();

        [Inject]
        private ITrailBalance_BreakupService _trailbalanceService { get; set; }
        [Inject]
        private IDBHelperService _idbHelperService { get; set; }
        [Inject]
        private IReportNavigationService ReportNavigationService { get; set; }
        [Inject]
        private INotificationService NotificationServiceManager { get; set; }
        [Inject]
        private IAuditService _auditService { get; set; }

        private List<TrailBalance_BreakupModel> getDBEntries { get; set; }
        private GroupsAccountsModel SelectedHead;
        private GroupsAccountsModel CurrentHead;
        private TrailBalance_BreakupModel CurrentAccount = new TrailBalance_BreakupModel();

        private Timer _timer;
        private DateTime FromDate { get; set; } = DateTime.Today.AddDays(-30);
        private DateTime ToDate { get; set; } = DateTime.Today;

        private bool IsLoading;
        private bool ShowDetail { get; set; } = true;
        List<TrailBalance_BreakupModel> statementmodelList = new();

        public void Dispose()
        {
            _timer?.Dispose();
        }
        protected override async Task OnInitializedAsync()
        {
            IsLoading = true;
            getDBEntries = null; // Clear previous data to show spinner
            StateHasChanged();

            try
            {
                // Pass !ShowDetail because VB6 logic implies "Summary" = Not Detail
                //getDBEntries = await ReportService.GenerateReport(FromDate, ToDate, !ShowDetail);
            }
            catch (Exception ex)
            {
                // You can use your NotificationService here
                Console.WriteLine(ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
        }
        private string GetRowStyle(TrailBalance_BreakupModel item)
        {
            // Match VB6 "SetMasterHeadColors" logic (Red & Bold for parents)
            if (item.Parent)
            {
                return "color: #b00000; background-color: #fff5f5;";
            }
            return "";
        }

        private void ViewLedger(ItemClickEventArgs e)
        {
            var item = e.Data as TrailBalance_BreakupModel;
            if (item != null)
            {
                // Navigate to Ledger page with AccNo param
                Navigation.NavigateTo($"/reports/ledger/{item.AccNo}");
            }
        }

        private void ResetForm()
        {
            StateHasChanged();
        }

        private void GoToIndexPage()
        {
            Navigation.NavigateTo("/", true);
        }
        private async Task Refreshlist()
        {
            IsLoading = true;
            try
            {
                var EntriesFromDb = await _trailbalanceService.GenerateReport(FromDate, ToDate, !ShowDetail);
                getDBEntries = EntriesFromDb.ToList();
                statementmodelList = getDBEntries;
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
        public async Task GenerateAndSaveReport(int itype)
        {
            if (statementmodelList != null && statementmodelList.Any())
            {   
                if (itype==1 || itype == 2 || itype == 3)
                { 
                    await _trailbalanceService.GenerateAndSaveReport(statementmodelList, false);
                }
                else
                {
                    await _trailbalanceService.GenerateAndSaveReport(statementmodelList,true);
                }

                string rptName = "";
                if (itype == 1)
                {
                    rptName = "TrialBalanceNew_Breakup.rpt";
                }
                else if (itype == 2)
                {
                    rptName = "TrialBalanceNewWMHV_Breakup.rpt";
                }
                else if (itype == 3)
                {
                    rptName = "TrialBalWCl_Breakup_Adjusted.rpt";
                }
                else if (itype == 4)
                {
                    rptName = "TrialBalWOp_Breakup.rpt"; 
                }
                else
                {
                    rptName = "TrialBalWCl_Breakup.rpt";
                }

                string strcond = "";

                //ReportNavigationService.PrintReportAsync(rptName, strcond);
                ReportNavigationService.PrintReportAsync_Old(rptName, strcond);
                StateHasChanged();
            }
        }

    }
}