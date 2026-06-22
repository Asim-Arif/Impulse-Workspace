using BlazorBootstrap;
using BlazorContextMenu;
using CIP.Pages;
using CIP.Services;
using DataAccessLibrary;
using DataAccessLibrary.DAC.Accounts;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models;
using DataAccessLibrary.Models.ViewModels;
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
    public partial class CashFlowMainHeads : ComponentBase, IDisposable
    {
        private CashFlowMainHeadsModel statementmodel = new();

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
        [Inject]
        private IVoucherService _voucherService { get; set; }

        private List<CashFlowMainHeadsModel> getDBEntries { get; set; } = new List<CashFlowMainHeadsModel>();
        private List<CashFlowType_DD_Model> CF_Types = new List<CashFlowType_DD_Model>();
        private List<CashFlowMain_DD_Model> CF_Heads = new List<CashFlowMain_DD_Model>();

        private CashFlowMainHeadsModel SelectedHead;
        private CashFlowMainHeadsModel CurrentHead;
        
        private CashFlowMain_DD_Model? SelectedMainHead = null;
        private CashFlowType_DD_Model? SelectedHeadType = null;

        private SubHeadAccountDetailModel? CurrentSubHead;
        private ChartOfAccountsModel? SelectedAccount = null;

        private Timer _timer;
        private bool ShowModal;
        private bool IsLoading = true;
        private bool IsLoadingDetail = true;
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
                var data = await _financialstatementService.GetCashFlowType_DD();
                CF_Types = data?.ToList() ?? new List<CashFlowType_DD_Model>();

                var data_heads = await _financialstatementService.GetCashFlowMainHeads_DD();
                CF_Heads = data_heads?.ToList() ?? new List<CashFlowMain_DD_Model>();

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
        private async Task SaveHead(int itype = 0)
        {
            if (IsAdding)
            {
                if (string.IsNullOrWhiteSpace(CurrentHead.Description))
                {
                    NotificationServiceManager.ShowWarning("Alert", "Cannot Save Without Main Group Name.");
                    return;
                }
                //CurrentHead.Type = SelectedHead.Type;

                CurrentHead.CFHM_RefID = SelectedMainHead.EntryID;
                CurrentHead.CFH_RefID = SelectedHeadType.EntryID ;

                await _financialstatementService.SaveCashFlowMainHeads(CurrentHead);
                NotificationServiceManager.ShowSuccess("Head Saved", "Saved successfully.");
            }
            else if (IsEdit)
            {
                if (string.IsNullOrWhiteSpace(CurrentHead.Description))
                {
                    NotificationServiceManager.ShowWarning("Alert", "Cannot Save Without Main Group Name.");
                    return;
                }
                await _financialstatementService.EditCashFlowMainHeads(CurrentHead, itype, CurrentHead.EntryID);
                NotificationServiceManager.ShowSuccess("Head Saved", "Saved successfully.");
            }

            ResetForm();
            HideModal();
            Refreshlist();
        }

        private async void EditSelected(ItemClickEventArgs e)
        {
            var bank = e.Data as CashFlowMainHeadsModel;
            SelectedHead = bank;

            if (SelectedHead != null)
            {
                CurrentHead = new CashFlowMainHeadsModel
                {
                    EntryID = SelectedHead.EntryID,
                    Description = SelectedHead.Description,
                    Notes = SelectedHead.Notes,
                    Marked = SelectedHead.Marked,
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
            if (CurrentHead != null)
            {
                if (!string.IsNullOrWhiteSpace(CurrentHead.Description))
                {
                    CurrentHead.Description = "";
                    CurrentHead.Notes = "";
                    CurrentHead.Marked =false;
                }
            }
            if (SelectedHead != null)
            {
                SelectedHead = null;
            }
            if (SelectedMainHead != null)
            {
                SelectedMainHead = null;
            }
            StateHasChanged();
            IsEdit = false;
            IsAdding = false;
            ShowModal = false;
        }

        private void GoToIndexPage()
        {
            Navigation.NavigateTo("/", true);
        }

        private void ShowAddModal()
        {
            CurrentHead = new CashFlowMainHeadsModel { };
            IsEdit = false;
            IsAdding = true;
            ShowModal = true;
        }

        private void HideModal()
        {
            CurrentHead = null;
            IsAdding = false;
            ShowModal = false;
            IsEdit = false;
        }
        private async Task Refreshlist()
        {
            IsLoading = true;
            try
            {
                var EntriesFromDb = await _financialstatementService.GetCashFlowMainHeadsList();
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
                var head = e.Data as CashFlowMainHeadsModel;
                SelectedHead = head;
                bool bdeleted = await _financialstatementService.DeletCashFlowMainHeads(SelectedHead.EntryID);
                if (bdeleted == false)
                {
                    NotificationServiceManager.ShowError("Delete Failed", "Could not delete the selected Head.");
                    return;
                }
                else
                {
                    NotificationServiceManager.ShowSuccess("Delete Alert", "Deleted Successfully.");
                    await Refreshlist();
                }
                StateHasChanged();
                ResetForm();
                Refreshlist();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("System Error Message", $"Failed to delete: {ex.Message}.");
            }
        }

        private async Task<IEnumerable<CashFlowMain_DD_Model>> SearchGetMainHeads(string searchText)
        {
            return await Task.FromResult(CF_Heads
                    .Where(x => x.CFMainHead != null &&
                                x.CFMainHead.ToLower().Contains(searchText.ToLower()))
                    .ToList());
        }
        private void SelectedResultChanged_MainHead(CashFlowMain_DD_Model? selectedhead)
        {
            SelectedMainHead = selectedhead;
        }

        private async Task<IEnumerable<CashFlowType_DD_Model>> SearchGetCashFlowTypes(string searchText)
        {
            return await Task.FromResult(CF_Types
                    .Where(x => x.CF_HeadType != null &&
                                x.CF_HeadType.ToLower().Contains(searchText.ToLower()))
                    .ToList());
        }
        private void SelectedResultChanged_types(CashFlowType_DD_Model? selectedhead)
        {
            SelectedHeadType = selectedhead;
        }

        private bool bShowAddDetailModel;
        private bool IsAddingSub;
        private bool IsEditSub;

        private ItemClickEventArgs _cachedMenuEvent;
        private int _ientrytype;
        private async Task ChangeMarkStatus(int itype,ItemClickEventArgs e)
        {
            IsEditSub = true;
            IsAddingSub = false;
            var head = e.Data as SubHeadAccountDetailModel;
            CurrentSubHead = head;
            if (CurrentSubHead != null)
            {
                _ientrytype = itype;
                IsEditSub =false;
            }
            StateHasChanged();
        }

    }
}