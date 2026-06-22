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
    public partial class DefineMainHeads : ComponentBase, IDisposable
    {
        private MainHeadsModel statementmodel = new();

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

        private List<MainHeadsModel> getDBEntries { get; set; } = new List<MainHeadsModel>();
        private MainHeadsModel SelectedHead;
        private MainHeadsModel CurrentHead;

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
                    NotificationServiceManager.ShowWarning("Alert", "Cannot Save Without Main Group Name.");
                    return;
                }
                await _financialstatementService.SaveMainHead(CurrentHead);
            }

            else if (IsEdit)
            {
                if (string.IsNullOrWhiteSpace(CurrentHead.GroupName))
                {
                    NotificationServiceManager.ShowWarning("Alert", "Cannot Save Without Main Group Name.");
                    return;
                }
                await _financialstatementService.EditMainHead(CurrentHead,CurrentHead.EntryID);
                NotificationServiceManager.ShowSuccess("Head Saved", "Main Group Saved successfully.");
            }
            ResetForm();
            HideModal();
            Refreshlist();
        }

        private async void EditSelectedBank(ItemClickEventArgs e)
        {
            var bank = e.Data as MainHeadsModel;
            SelectedHead = bank;

            if (SelectedHead != null)
            {
                CurrentHead = new MainHeadsModel
                {
                    EntryID = SelectedHead.EntryID,
                    GroupName= SelectedHead.GroupName,
                    SortNo = SelectedHead.SortNo,
                    Description = SelectedHead.Description,
                };

                IsAdding = false;
                ShowModal = true;
                IsEdit = true;
                StateHasChanged();
            }
        }

        private void ResetForm()
        {
            if (CurrentHead!=null)
            { 
                if (!string.IsNullOrWhiteSpace(CurrentHead.GroupName))
                { 
                    CurrentHead.GroupName= "";
                    CurrentHead.SortNo= "";
                    CurrentHead.Description = "";
                }
            }
            StateHasChanged();
        }

        private void GoToIndexPage()
        {
            Navigation.NavigateTo("/", true);
        }

        private void ShowAddModal()
        {
            CurrentHead = new MainHeadsModel { };
            IsAdding = true;
            ShowModal = true;
            IsEdit = false;

            if (CurrentHead!= null)
            { 
                long sortno = _idbHelperService.getSingleIntValue("MAX(SortNo)", "GroupsForFinancialReport");
                sortno = sortno + 1;
                CurrentHead.SortNo = sortno.ToString();
            }
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
                var EntriesFromDb = await _financialstatementService.GetMainHeadsList();
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
                var head = e.Data as MainHeadsModel;
                SelectedHead = head;
                bool bdeleted= await _financialstatementService.DeleteMainHead(SelectedHead.EntryID);
                if (bdeleted == false)
                {
                    NotificationServiceManager.ShowError("Delete Failed", "Could not delete the selected Main Head.");
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
    }
}