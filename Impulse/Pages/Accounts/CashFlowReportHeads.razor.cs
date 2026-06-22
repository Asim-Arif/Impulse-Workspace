using BlazorBootstrap;
using BlazorContextMenu;
using CIP.Pages;
using CIP.Services;
using DataAccessLibrary;
using DataAccessLibrary.DAC.Accounts;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Mvc.Rendering;
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


namespace CIP.Pages.Accounts
{
    public partial class CashFlowReportHeads : ComponentBase, IDisposable
    {
        private SubHeadsManualEntryModel statementmodel = new();

        [Inject]
        private IFinancialStatementService _financialstatementService { get; set; }
        [Inject]
        private IDBHelperService _idbHelperService { get; set; }
        [Inject]
        private IReportNavigationService ReportNavigationService { get; set; }
        [Inject]
        private INotificationService NotificationServiceManager { get; set; }
        [Inject]
        private IAuditService _auditService { get; set; }

        private List<CashFlowReportHeadsDetailToLoadModel> getDBEntries { get; set; } = new List<CashFlowReportHeadsDetailToLoadModel>();
        private CashFlowReportHeads_Main_Model SelectedHead;
        private CashFlowReportHeadsDetailToLoadModel SelectedEntry;
        private CashFlowReportHeads_Main_Model CurrentHead =new CashFlowReportHeads_Main_Model();

        private SubHeadsListModel? SelectedSubHead = null;
        //private MainHeadsListModel? SelectedMainHead = null;
        private List<SubHeadsListModel> SubHeads = new List<SubHeadsListModel>();

        //private List<MainHeadsListModel> MainHeads = new List<MainHeadsListModel>();
        public IList<CashFlowHeadsListModel> MainHeads { get; set; } = new List<CashFlowHeadsListModel>();
        private IList<CashFlowHeadsListModel> SelectedMainHead { get; set; } = new List<CashFlowHeadsListModel>();

        private Timer _timer;
        private bool ShowModal;
        private bool IsLoading=true;
        private bool IsAdding;
        private bool IsEdit;
        private bool bshowinactive = false;

        public void Dispose()
        {
            _timer?.Dispose();
        }
        protected override async Task OnInitializedAsync()
        {
            IsLoading = true;
            try
            {
                var data = await _financialstatementService.GetCashFlowHeadsList_DD();
                MainHeads = data?.ToList() ?? new List<CashFlowHeadsListModel>();

                await Refreshlist();
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
        private async Task SaveHead()
        {
            if (IsAdding)
            {
                if (CurrentHead.Description == null)
                {
                    NotificationServiceManager.ShowWarning("Alert", "Cannot Save Without Description.");
                    return;
                }
                else if (SelectedMainHead.Count == 0)
                {
                    NotificationServiceManager.ShowWarning("Alert", "Cannot Save Without Accounts Groups.");
                    return;
                }

                CurrentHead.DetailItems = SelectedMainHead.Select(x => new CashFlowReportHeads_Detail_Model
                {
                    Group_RefID = x.EntryID
                }).ToList();

                await _financialstatementService.SaveCashFlowReportHeads(CurrentHead);
                SelectedMainHead = null;
            }

            else if (IsEdit)
            {
                if (string.IsNullOrWhiteSpace(CurrentHead.Description))
                {
                    NotificationServiceManager.ShowWarning("Alert", "Cannot Save Without Description.");
                    return;
                }
                await _financialstatementService.EditCashFlowReportHeads(CurrentHead,CurrentHead.EntryID);
                NotificationServiceManager.ShowSuccess("Head Saved", "Saved successfully.");
            }
            ResetForm();
            HideModal();
            Refreshlist();
        }

        private async void EditSelected(ItemClickEventArgs e)
        {
            var main = e.Data as CashFlowReportHeadsDetailToLoadModel;
            SelectedEntry = main;

            if (SelectedEntry != null)
            {
                CurrentHead = new CashFlowReportHeads_Main_Model
                {
                    EntryID = SelectedEntry.RefID,
                    Description = SelectedEntry.Description,
                    NoteNo = SelectedEntry.NoteNo,
                    Marked = SelectedEntry.Marked,
                    IsTotal = SelectedEntry.IsTotal,
                    SrNo = SelectedEntry.SrNo
                };

                IsAdding = false;
                ShowModal = true;
                IsEdit = true;
                StateHasChanged();
            }
        }

        private void ResetForm()
        {
            StateHasChanged();
            if (!string.IsNullOrWhiteSpace(CurrentHead.Description))
            { 
                CurrentHead.Description= "";
                CurrentHead.NoteNo= "";
                CurrentHead.Description = "";
            }
            if (SelectedSubHead != null) 
            {
                SelectedSubHead.GroupName=string.Empty;
                SelectedSubHead.EntryID = 0;
            }

            StateHasChanged();
        }

        private void GoToIndexPage()
        {
            Navigation.NavigateTo("/", true);
        }

        private void ShowAddModal()
        {
            CurrentHead = new CashFlowReportHeads_Main_Model { };
            CurrentHead.SrNo = _idbHelperService.getSingleIntValue("COUNT(*)", "vProfitnLossHeads_Detail")+1;
            IsAdding = true;
            ShowModal = true;
            IsEdit = false;
        }
        private void HideModal()
        {
            CurrentHead = null;
            IsAdding = false;
            ShowModal = false;
            IsEdit = false;

            if (SelectedSubHead != null)
            {
                SelectedSubHead.GroupName = string.Empty;
                SelectedSubHead.EntryID = 0;
            }

        }
        private async Task Refreshlist()
        {
            IsLoading = true;
            try
            {
                var EntriesFromDb = await _financialstatementService.GetCashFlowReportHeadsList();
                getDBEntries = EntriesFromDb.ToList();
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
        private async Task DeleteSelected(ItemClickEventArgs e)
        {
            try
            {

                var head = e.Data as CashFlowReportHeadsDetailToLoadModel;
                SelectedEntry = head;    
                bool bdeleted= await _financialstatementService.DeleteCashFlowReportHeads(SelectedEntry.RefID);
                if (bdeleted == false)
                {
                    NotificationServiceManager.ShowError("Delete Failed", "Could not delete the selected Head.");
                    return;
                }
                else
                {
                    NotificationServiceManager.ShowSuccess("Delete Alert", "Deleted Successfully.");
                }
                StateHasChanged();
                Refreshlist();
            }
            catch (Exception ex)
            { 
                NotificationServiceManager.ShowError("System Error Message", $"Failed to delete: {ex.Message}.");
            }
        }

        private async Task<IEnumerable<SubHeadsListModel>> SearchGetSubHeads(string searchText)
        {
            return await Task.FromResult(SubHeads
                    .Where(x => x.GroupName != null &&
                                x.GroupName.ToLower().Contains(searchText.ToLower()))
                    .ToList());
        }
        private void SelectedResultChanged_SubHead(SubHeadsListModel? selectedhead)
        {
            SelectedSubHead = selectedhead;
        }

        private async Task<IEnumerable<CashFlowHeadsListModel>> SearchMainHeads(string searchText)
        {
            return await Task.FromResult(MainHeads
                    .Where(x => x.GroupName != null &&
                                x.GroupName.ToLower().Contains(searchText.ToLower()))
                    .ToList());
        }
        private async Task SelectedResultChanged_MainHead(IList<CashFlowHeadsListModel?> selectedhead)
        {
            //SelectedMainHead = selectedhead;
            SelectedMainHead = selectedhead ?? new List<CashFlowHeadsListModel>();
            if (SelectedMainHead != null)
            {
                Console.WriteLine($"You have selected {SelectedMainHead.Count} items.");
                await Task.CompletedTask;
            }
        }

    }
}