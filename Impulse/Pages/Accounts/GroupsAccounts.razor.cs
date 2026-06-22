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
    public partial class GroupsAccounts : ComponentBase, IDisposable
    {
        private GroupsAccountsModel statementmodel = new();

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

        private List<GroupsAccountsModel> getDBEntries { get; set; } = new List<GroupsAccountsModel>();
        private GroupsAccountsModel SelectedHead;
        private GroupsAccountsModel CurrentHead;

        private Timer _timer;
        private bool ShowModal;
        private bool IsLoading=true;
        private bool IsAdding;
        private bool IsEdit;
        private bool bshowinactive = false;

        public string Type { get; set; } = "Expense";

        public void Dispose()
        {
            _timer?.Dispose();
        }
        protected override async Task OnInitializedAsync()
        {
            IsLoading = true;
            try
            {
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
                if (string.IsNullOrWhiteSpace(CurrentHead.GroupName))
                {
                    NotificationServiceManager.ShowWarning("Alert", "Cannot Save Without Group Name.");
                    return;
                }
                else if (string.IsNullOrWhiteSpace(CurrentHead.Type))
                {
                    NotificationServiceManager.ShowWarning("Alert", "Cannot Save Without Group Type.");
                    return;
                }
                else if (string.IsNullOrWhiteSpace(CurrentHead.MainGroupName))
                {
                    NotificationServiceManager.ShowWarning("Alert", "Cannot Save Without Main Group Name.");
                    return;
                }
                await _financialstatementService.SaveGroupsAccounts(CurrentHead);
            }

            else if (IsEdit)
            {
                if (string.IsNullOrWhiteSpace(CurrentHead.GroupName))
                {
                    NotificationServiceManager.ShowWarning("Alert", "Cannot Save Without Group Name.");
                    return;
                }
                else if (string.IsNullOrWhiteSpace(CurrentHead.Type))
                {
                    NotificationServiceManager.ShowWarning("Alert", "Cannot Save Without Group Type.");
                    return;
                }
                else if (string.IsNullOrWhiteSpace(CurrentHead.MainGroupName))
                {
                    NotificationServiceManager.ShowWarning("Alert", "Cannot Save Without Main Group Name.");
                    return;
                }
                await _financialstatementService.EditGroupsAccounts(CurrentHead,CurrentHead.EntryID);
                NotificationServiceManager.ShowSuccess("Head Saved", "Main Group Saved successfully.");
            }
            ResetForm();
            HideModal();
            await Refreshlist();
        }

        private async void EditSelectedBank(ItemClickEventArgs e)
        {
            var entry = e.Data as GroupsAccountsModel;
            SelectedHead = entry;

            if (SelectedHead != null)
            {
                CurrentHead = new GroupsAccountsModel
                {
                    EntryID = SelectedHead.EntryID,
                    GroupName= SelectedHead.GroupName,
                    SortNo = SelectedHead.SortNo,
                    Type = SelectedHead.Type,
                    MainGroupName = SelectedHead.MainGroupName
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
                CurrentHead.GroupName= "";
                CurrentHead.SortNo= "";
                CurrentHead.Type = "";
                CurrentHead.MainGroupName = "";
            }
            StateHasChanged();
        }

        private void GoToIndexPage()
        {
            Navigation.NavigateTo("/", true);
        }

        private void ShowAddModal()
        {
            CurrentHead = new GroupsAccountsModel { };
            CurrentHead.Type = Type;
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
        }
        private async Task Refreshlist()
        {
            IsLoading = true;
            try
            {
                var EntriesFromDb = await _financialstatementService.GetGroupsAccountsList();
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
                var head = e.Data as GroupsAccountsModel;
                SelectedHead = head;
                bool bdeleted= await _financialstatementService.DeleteGroupsAccounts(SelectedHead.EntryID);
                if (bdeleted == false)
                {
                    NotificationServiceManager.ShowError("Delete Failed", "Could not delete the selected Group.");
                    return;
                }
                else
                {
                    NotificationServiceManager.ShowSuccess("Delete Alert", "Deleted Successfully.");
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
    }
}