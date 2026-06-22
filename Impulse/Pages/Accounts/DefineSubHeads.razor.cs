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
    public partial class DefineSubHeads : ComponentBase, IDisposable
    {
        private SubHeadsModel statementmodel = new();

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

        private List<SubHeadsModel> getDBEntries { get; set; } = new List<SubHeadsModel>();
        private List<SubHeadsModel> MainHeads = new List<SubHeadsModel>();
        private List<ChartOfAccountsModel> Accounts = new List<ChartOfAccountsModel>();
        private List<SubHeadAccountDetailModel> getDBEntries_accounthead { get; set; } = new List<SubHeadAccountDetailModel>();

        private SubHeadsModel SelectedHead;
        private SubHeadsModel CurrentHead;
        private SubHeadsModel? SelectedMainHead = null;
        private SubHeadAccountDetailModel? SelectedSubHead;
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
        private async Task SaveHead(int itype = 0)
        {
            if (IsAdding)
            {
                if (string.IsNullOrWhiteSpace(CurrentHead.GroupName))
                {
                    NotificationServiceManager.ShowWarning("Alert", "Cannot Save Without Main Group Name.");
                    return;
                }
                CurrentHead.Type = SelectedHead.Type;
                await _financialstatementService.SaveSubHead(CurrentHead);
                NotificationServiceManager.ShowSuccess("Head Saved", "Saved successfully.");
            }
            else if (IsEdit)
            {
                if (string.IsNullOrWhiteSpace(CurrentHead.GroupName))
                {
                    NotificationServiceManager.ShowWarning("Alert", "Cannot Save Without Main Group Name.");
                    return;
                }
                await _financialstatementService.EditSubHead(CurrentHead, CurrentHead.EntryID, itype);
                NotificationServiceManager.ShowSuccess("Head Saved", "Saved successfully.");
            }
            else
            {
                await _financialstatementService.EditSubHead(SelectedHead, SelectedHead.EntryID, itype);
            }
            ResetForm();
            HideModal();
            Refreshlist();
        }

        private async void EditSelected(ItemClickEventArgs e)
        {
            var bank = e.Data as SubHeadsModel;
            SelectedHead = bank;

            if (SelectedHead != null)
            {
                CurrentHead = new SubHeadsModel
                {
                    EntryID = SelectedHead.EntryID,
                    GroupName = SelectedHead.GroupName,
                    SubDesc = SelectedHead.SubDesc,
                    Notes = SelectedHead.Notes,
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
                if (!string.IsNullOrWhiteSpace(CurrentHead.GroupName))
                {
                    CurrentHead.GroupName = "";
                    CurrentHead.Notes = "";
                    CurrentHead.SubDesc = "";
                    CurrentHead.ShowDetail = false;
                    CurrentHead.bShowSubHeadDetail_NTA = false;
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
            CurrentHead = new SubHeadsModel { };
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
                var EntriesFromDb = await _financialstatementService.GetSubHeadList();
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
                var head = e.Data as SubHeadsModel;
                SelectedHead = head;
                bool bdeleted = await _financialstatementService.DeleteSubHead(SelectedHead.EntryID);
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

        private async Task<IEnumerable<SubHeadsModel>> SearchGetMainHeads(string searchText)
        {
            return await Task.FromResult(MainHeads
                    .Where(x => x.MainGroupName != null &&
                                x.MainGroupName.ToLower().Contains(searchText.ToLower()))
                    .ToList());
        }
        private void SelectedResultChanged_MainHead(SubHeadsModel? selectedhead)
        {
            SelectedMainHead = selectedhead;
        }

        private async Task<IEnumerable<ChartOfAccountsModel>> SearchSelectedAccount(string searchText)
        {
            return await Task.FromResult(Accounts
                    .Where(x => x.AccTitle != null &&
                                x.AccTitle.ToLower().Contains(searchText.ToLower()))
                    .ToList());
        }
        private void SelectedResultChanged_Accounts(ChartOfAccountsModel? selectedaccount)
        {
            SelectedAccount = selectedaccount;
        }

        private bool bShowAddDetailModel;
        private bool IsAddingSub;
        private bool IsEditSub;
        private async Task ShowAddDetailModel(ItemClickEventArgs e)
        {
            var selected = e.Data as SubHeadsModel;
            SelectedHead = selected;
            IsAddingSub = true;
            IsEditSub = false;
            IsLoadingDetail = true;
            if (SelectedHead != null)
            {
                SelectedSubHead = new SubHeadAccountDetailModel();
                bShowAddDetailModel = true;
                Accounts = await _voucherService.GetTransactionalAccounts();
                await RefreshSubHeadsAccountslist();
                StateHasChanged();
            }
        }

        private void HideAddDetailModal()
        {
            SelectedSubHead = null;
            bShowAddDetailModel = false;
            IsAddingSub = true;
            ShowModal = false;
            IsEditSub = false;
            IsLoadingDetail = false;
        }
        private async Task RefreshSubHeadsAccountslist()
        {
            IsLoadingDetail = true;
            try
            {
                if (SelectedHead !=null) { 
                    var EntriesFromDb = await _financialstatementService.GetAccountGroup_HeadsDetail(SelectedHead.EntryID);
                    getDBEntries_accounthead = EntriesFromDb.ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching Data: {ex.Message}");
            }
            finally
            {
                IsLoadingDetail = false;
                StateHasChanged();
            }
        }

        private async Task SaveSubHeadEntry()
        {
            if (IsAddingSub)
            {
                if (string.IsNullOrWhiteSpace(SelectedSubHead.Description))
                {
                    NotificationServiceManager.ShowWarning("Alert", "Cannot Save Without Main Group Name.");
                    return;
                }

                SelectedSubHead.AccGroup_RefID = SelectedHead.EntryID;
                SelectedSubHead.AccNo = SelectedAccount.AccNo;
                SelectedSubHead.UserName =  _auditService.GetCurrentUserName();
                SelectedSubHead.MachineName = _auditService.GetClientIpAddress();

                await _financialstatementService.SaveSubHeadDetail(SelectedSubHead);
                NotificationServiceManager.ShowSuccess("Saved Alert", "Saved successfully.");

                ResetAccountModel();
            }
            else if (IsEditSub)
            {
                SelectedSubHead = CurrentSubHead;

                if (string.IsNullOrWhiteSpace(SelectedSubHead.Description))
                {
                    NotificationServiceManager.ShowWarning("Saved Alert", "Cannot Save Without Description.");
                    return;
                }
                await _financialstatementService.EditSubHeadDetailEntry(SelectedSubHead, SelectedSubHead.EntryID, _ientrytype);
                NotificationServiceManager.ShowSuccess("Saved Alert", "Saved successfully.");

                SelectedSubHead.NoteNo = string.Empty;
                SelectedSubHead.Description = string.Empty;
                CurrentSubHead.Notes= string.Empty;
                CurrentSubHead.Description = string.Empty;

                _ientrytype = 0;
                HideSubHeadEntryModel();
                StateHasChanged();
            }
            RefreshSubHeadsAccountslist();
            ResetAccountModel();

        }
        private async void EditSelectedSubHeadEntry(ItemClickEventArgs e)
        {
            var subhead = e.Data as SubHeadAccountDetailModel;
            CurrentSubHead = subhead;

            if (CurrentSubHead != null)
            {
                CurrentSubHead = new SubHeadAccountDetailModel
                {
                    EntryID = CurrentSubHead.EntryID,
                    Description = CurrentSubHead.Description,
                    NoteNo = CurrentSubHead.NoteNo,
                    AccNo = CurrentSubHead.AccNo,
                };

                IsAddingSub = false;
                IsEditSub = true;
                StateHasChanged();
                RefreshSubHeadsAccountslist();
            }
        }

        private async Task DeleteSelectedSubHeadEntry(ItemClickEventArgs e)
        {
            try
            {
                var head = e.Data as SubHeadAccountDetailModel;
                SelectedSubHead = head;
                bool bdeleted = await _financialstatementService.DeleteSubHeadEntry(SelectedSubHead.EntryID);
                if (bdeleted == false)
                {
                    NotificationServiceManager.ShowError("Delete Failed", "Could not delete the selected Head.");
                    return;
                }
                else
                {
                    NotificationServiceManager.ShowSuccess("Delete Alert", "Deleted Successfully.");
                    await RefreshSubHeadsAccountslist();
                }
                StateHasChanged();
                ResetForm();
                RefreshSubHeadsAccountslist();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("System Error Message", $"Failed to delete: {ex.Message}.");
            }
        }

        private void HideSubHeadEntryModel()
        {
            //SelectedSubHead = null;
            CurrentSubHead = null;
            IsAddingSub =true;
            ShowModal = false;
            IsEditSub = false;
            IsLoadingDetail = false;
            StateHasChanged();
        }

        private void ResetAccountModel()
        {
            SelectedSubHead.NoteNo = "";
            SelectedSubHead.Description = "";
            if (SelectedAccount != null)
            {
                SelectedAccount = null;
            }
            StateHasChanged();
        }

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
                SaveSubHeadEntry();
                IsEditSub =false;
            }
            StateHasChanged();
        }

    }
}