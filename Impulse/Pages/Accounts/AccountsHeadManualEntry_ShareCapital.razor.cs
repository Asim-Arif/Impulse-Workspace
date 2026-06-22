using BlazorBootstrap;
using BlazorContextMenu;
using DataAccessLibrary;
using DataAccessLibrary.DAC.Accounts;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
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
    public partial class AccountsHeadManualEntry_ShareCapital : ComponentBase, IDisposable
    {
        private AccountsHeadManualEntry_SC_Model statementmodel = new();

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

        private List<AccountsHeadManualEntry_SC_Model> getDBEntries { get; set; } = new List<AccountsHeadManualEntry_SC_Model>();
        private AccountsHeadManualEntry_SC_Model SelectedHead;
        private AccountsHeadManualEntry_SC_Model CurrentHead =new AccountsHeadManualEntry_SC_Model();

        private SubHeadsListModel? SelectedSubHead = null;
        private List<SubHeadsListModel> SubHeads = new List<SubHeadsListModel>();

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
                CurrentHead.DT = DateTime.Now;

                var data = await _financialstatementService.GetSubHeads(0);
                SubHeads = data?.ToList() ?? new List<SubHeadsListModel>();

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
                else if (CurrentHead.TotalAmount == 0)
                {
                    NotificationServiceManager.ShowWarning("Alert", "Invalid Amount.");
                    return;
                }
                CurrentHead.AccGroup_RefID = SelectedSubHead.EntryID;
                await _financialstatementService.SaveAccountsHeadManualEntries_SC(CurrentHead);
            }

            else if (IsEdit)
            {
                if (string.IsNullOrWhiteSpace(CurrentHead.GroupName))
                {
                    NotificationServiceManager.ShowWarning("Alert", "Cannot Save Without Main Group Name.");
                    return;
                }
                await _financialstatementService.EditAccountsHeadManualEntries_SC(CurrentHead,CurrentHead.EntryID);
                NotificationServiceManager.ShowSuccess("Head Saved", "Main Group Saved successfully.");
            }
            ResetForm();
            HideModal();
            Refreshlist();
        }

        private async void EditSelectedBank(ItemClickEventArgs e)
        {
            var bank = e.Data as AccountsHeadManualEntry_SC_Model;
            SelectedHead = bank;

            if (SelectedHead != null)
            {
                CurrentHead = new AccountsHeadManualEntry_SC_Model
                {
                    EntryID = SelectedHead.EntryID,
                    GroupName= SelectedHead.GroupName,
                    Notes = SelectedHead.Notes,
                    Description = SelectedHead.Description,
                    TotalAmount = SelectedHead.TotalAmount,
                    DT = SelectedHead.DT,
                    AccGroup_RefID= SelectedHead.AccGroup_RefID,
                    ShareValue = SelectedHead.ShareValue,
                    ShareRate = SelectedHead.ShareRate
                };

                SelectedSubHead = new SubHeadsListModel
                {
                    EntryID = SelectedHead.AccGroup_RefID,
                    GroupName = SelectedHead.GroupName
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
            }
            StateHasChanged();
        }

        private void GoToIndexPage()
        {
            Navigation.NavigateTo("/", true);
        }

        private void ShowAddModal()
        {
            CurrentHead = new AccountsHeadManualEntry_SC_Model { };
            CurrentHead.DT= DateTime.Now;
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
                var EntriesFromDb = await _financialstatementService.GetAccountsHeadManualEntries_SC();
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
                var head = e.Data as AccountsHeadManualEntry_SC_Model;
                SelectedHead = head;    
                bool bdeleted= await _financialstatementService.DeleteAccountsHeadManualEntries_SC(SelectedHead.EntryID);
                if (bdeleted == false)
                {
                    NotificationServiceManager.ShowError("Delete Failed", "Could not delete the selected Main Head.");
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
        private void CalculateTotalAmount() 
        { 
            double totalAmount = 0;
            if (CurrentHead.ShareValue > 0 && CurrentHead.ShareRate > 0)
            {
                totalAmount = CurrentHead.ShareValue * CurrentHead.ShareRate;
                CurrentHead.TotalAmount = totalAmount;
            }
        }

    }
}