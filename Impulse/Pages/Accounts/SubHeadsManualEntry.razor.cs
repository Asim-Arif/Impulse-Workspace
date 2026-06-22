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


namespace CIP.Pages.Accounts
{
    public partial class SubHeadsManualEntry : ComponentBase, IDisposable
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

        private List<SubHeadsManualEntryModel> getDBEntries { get; set; } = new List<SubHeadsManualEntryModel>();
        private SubHeadsManualEntryModel SelectedHead;
        private SubHeadsManualEntryModel CurrentHead =new SubHeadsManualEntryModel();

        private SubHeadsListModel? SelectedSubHead = null;
        private SubHeadsModel? SelectedMainHead = null;
        private List<SubHeadsListModel> SubHeads = new List<SubHeadsListModel>();
        
        private List<SubHeadsModel> MainHeads = new List<SubHeadsModel>();

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
                CurrentHead.DTFrom = DateTime.Now;
                CurrentHead.DTTo = DateTime.Now;

                //var data = await _financialstatementService.GetSubHeads();
                //SubHeads = data?.ToList() ?? new List<SubHeadsListModel>();

                var data = await _financialstatementService.GetMainHeads();
                MainHeads = data?.ToList() ?? new List<SubHeadsModel>();

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
                if (SelectedSubHead == null)
                {
                    NotificationServiceManager.ShowWarning("Alert", "Cannot Save Sub Head Name.");
                    return;
                }
                CurrentHead.AccGroup_RefID = SelectedSubHead.EntryID;
                await _financialstatementService.SaveSubHeadsManualEntries(CurrentHead);
            }

            else if (IsEdit)
            {
                if (string.IsNullOrWhiteSpace(CurrentHead.GroupName))
                {
                    NotificationServiceManager.ShowWarning("Alert", "Cannot Save Without Main Group Name.");
                    return;
                }
                await _financialstatementService.EditSubHeadsManualEntries(CurrentHead,CurrentHead.EntryID);
                NotificationServiceManager.ShowSuccess("Head Saved", "Main Group Saved successfully.");
            }
            ResetForm();
            HideModal();
            Refreshlist();
        }

        private async void EditSelectedBank(ItemClickEventArgs e)
        {
            var bank = e.Data as SubHeadsManualEntryModel;
            SelectedHead = bank;

            if (SelectedHead != null)
            {
                CurrentHead = new SubHeadsManualEntryModel
                {
                    EntryID = SelectedHead.EntryID,
                    GroupName= SelectedHead.GroupName,
                    Notes = SelectedHead.Notes,
                    Description = SelectedHead.Description,
                    OpeningBalance = SelectedHead.OpeningBalance,
                    ClosingBalance = SelectedHead.ClosingBalance,
                    DTFrom = SelectedHead.DTFrom,
                    DTTo = SelectedHead.DTTo,
                    AccGroup_RefID = SelectedHead.AccGroup_RefID,
                    Marked = SelectedHead.Marked,
                    SrNo = SelectedHead.SrNo,
                    MainAccGroup_RefID = SelectedHead.MainAccGroup_RefID,
                    MainGroupName = SelectedHead.MainGroupName,
                    MainNoteNo = SelectedHead.MainNoteNo,
                };

                SelectedSubHead = new SubHeadsListModel
                {
                    EntryID = SelectedHead.AccGroup_RefID,
                    GroupName = SelectedHead.GroupName
                };

                SelectedMainHead = new SubHeadsModel 
                {
                    EntryID = SelectedHead.MainAccGroup_RefID,
                    MainGroupName = SelectedHead.MainGroupName,
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
            if (!string.IsNullOrWhiteSpace(CurrentHead.GroupName))
            { 
                CurrentHead.GroupName= "";
                CurrentHead.Notes= "";
                CurrentHead.Description = "";
                CurrentHead.AccGroup_RefID = 0;
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
            CurrentHead = new SubHeadsManualEntryModel { };
            CurrentHead.DTFrom= DateTime.Now;
            CurrentHead.DTTo = DateTime.Now;
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

            if (SelectedMainHead != null)
            {
                SelectedMainHead.MainGroupName= string.Empty;
                SelectedMainHead.Type = 0;
            }

        }
        private async Task Refreshlist()
        {
            IsLoading = true;
            try
            {
                var EntriesFromDb = await _financialstatementService.GetSubHeadsManualEntriesList();
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

                var head = e.Data as SubHeadsManualEntryModel;
                SelectedHead = head;    
                bool bdeleted= await _financialstatementService.DeleteSubHeadsManualEntry(SelectedHead.EntryID);
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

        private async Task<IEnumerable<SubHeadsModel>> SearchMainHeads(string searchText)
        {
            return await Task.FromResult(MainHeads
                    .Where(x => x.MainGroupName != null &&
                                x.MainGroupName.ToLower().Contains(searchText.ToLower()))
                    .ToList());
        }
        //private void SelectedResultChanged_MainHead(SubHeadsModel? selectedhead)
        private async Task SelectedResultChanged_MainHead(SubHeadsModel? selectedhead)
        {
            SelectedMainHead = selectedhead;

            if (SelectedMainHead != null) 
            {
                var data = await _financialstatementService.GetSubHeads(SelectedMainHead.Type);
                SubHeads = data?.ToList() ?? new List<SubHeadsListModel>();
            }

            if (SelectedSubHead != null)
            {
                SelectedSubHead.GroupName = string.Empty;
                SelectedSubHead.EntryID = 0;
            }
        }

    }
}