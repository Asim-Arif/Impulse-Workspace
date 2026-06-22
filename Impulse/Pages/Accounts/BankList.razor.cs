using BlazorBootstrap;
using BlazorContextMenu;
using DataAccessLibrary;
using DataAccessLibrary.DAC.Accounts;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DocumentFormat.OpenXml.Office2010.CustomUI;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Microsoft.JSInterop;
using Microsoft.VisualBasic;
using Radzen;
using System.Collections.ObjectModel;
using System.Security.AccessControl;
using System.Threading.Tasks;
using OpenXmlMenu = DocumentFormat.OpenXml.Office2010.CustomUI;

namespace Impulse.Pages.Accounts
{
    public partial class BankList : ComponentBase
    {
        private BankListModel banklistmodel = new();

        [Inject]
        protected IBankListDataAccess BankListDataAccess { get; set; }
        [Inject]
        protected IDBHelper IDBHelper { get; set; }
        [Inject]
        protected IChartOfAccountsDataAccess iChartOfAccountsDataAccess { get; set; }
        [Inject]
        private INotificationService NotificationServiceManager { get; set; }

        private List<BankListModel> BankListfromDB { get; set; } = new List<BankListModel>();
        private BankListModel SelectedBank;
        private BankListModel CurrentBank = new BankListModel();
        private List<BankListModel> accheadsforlist = new List<BankListModel>();

        private bool ShowModal;
        private bool IsLoading=true;
        private bool IsAdding;
        private bool IsEdit;
        private bool bshowinactive = false;

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
        private async Task SaveBank()
        {
            if (IsAdding)
            {
                if (string.IsNullOrWhiteSpace(CurrentBank.Bank))
                {
                    await JS.InvokeVoidAsync("alert", "Cannot Save Bank. Cannot Save Without Bank Name.");
                    return;
                }
                else if (string.IsNullOrWhiteSpace(CurrentBank.Branch))
                {
                    await JS.InvokeVoidAsync("alert", "Cannot Save Bank. Cannot Save Without Branch Name.");
                    return;
                }
                else if (string.IsNullOrWhiteSpace(CurrentBank.City))
                {
                    await JS.InvokeVoidAsync("alert", "Cannot Save Bank. Cannot Save Without City");
                    return;
                }
                if (CurrentBank.COA==true)
                {
                    string strBankParent = await IDBHelper.getSingleStringValue("DataValue", "GeneralData", "WHERE DataName='BankParent'");
                    string NextAccountNo = await iChartOfAccountsDataAccess.GetNextAccountNumberAsync("22", CurrentBank.Bank, false, strBankParent);
                    CurrentBank.AccNo = NextAccountNo;
                    CurrentBank.SubAccOf= strBankParent;
                    CurrentBank.AccType = "22";
                }
                await BankListDataAccess.SaveNewBank(CurrentBank);
            }

            else if (IsEdit)
            {
                if (string.IsNullOrWhiteSpace(CurrentBank.Bank))
                {
                    await JS.InvokeVoidAsync("alert", "Cannot Save Bank. Cannot Save Without Bank Name.");
                    return;
                }
                else if (string.IsNullOrWhiteSpace(CurrentBank.Branch))
                {
                    await JS.InvokeVoidAsync("alert", "Cannot Save Bank. Cannot Save Without Branch Name.");
                    return;
                }
                else if (string.IsNullOrWhiteSpace(CurrentBank.City))
                {
                    await JS.InvokeVoidAsync("alert", "Cannot Save Bank. Cannot Save Without City.");
                    return;
                }
                await BankListDataAccess.EditBankAsync(CurrentBank, CurrentBank.BankID);
                ShowSuccess("Bank Saved", "Bank Saved successfully.");
            }
            ResetForm();
            HideModal();
            Refreshlist();
        }

        private async void EditSelectedBank(ItemClickEventArgs e)
        {
            var bank = e.Data as BankListModel;
            SelectedBank = bank;

            if (SelectedBank != null)
            {
                CurrentBank = new BankListModel
                {
                    BankID = SelectedBank.BankID,
                    Bank= SelectedBank.Bank,
                    Branch = SelectedBank.Branch,
                    City = SelectedBank.City,
                    Phone1 = SelectedBank.Phone1,
                    Phone2 = SelectedBank.Phone2,
                    FaxNo = SelectedBank.FaxNo,
                    Remittance = SelectedBank.Remittance,
                    Manager = SelectedBank.Manager,
                    Address = SelectedBank.Address,
                    ShowInCIP = SelectedBank.ShowInCIP,
                    COA = SelectedBank.COA,
                    BankAccNo=SelectedBank.BankAccNo,
                    Email = SelectedBank.Email,
                };

                IsAdding = false;
                ShowModal = true;
                IsEdit = true;
                StateHasChanged();
                Console.WriteLine($"Editing: {SelectedBank.BankID}");
            }
        }

        private void ResetForm()
        {
            StateHasChanged();
            if (!string.IsNullOrWhiteSpace(CurrentBank.Bank))
            { 
                CurrentBank.Bank= "";
                CurrentBank.BankAccNo= "";
                CurrentBank.Branch = "";
                CurrentBank.City = "";
                CurrentBank.Email = "";
                CurrentBank.FaxNo = "";
                CurrentBank.Phone1 = "";
                CurrentBank.Phone2 = "";
                CurrentBank.Address = "";
                CurrentBank.ShowInCIP = false;
                CurrentBank.BankID = 0;
                SelectedBank = null;
            }
            StateHasChanged();
        }

        private void GoToIndexPage()
        {
            Navigation.NavigateTo("/", true);
        }

        private void ShowAddModal()
        {
            CurrentBank = new BankListModel { };
            IsAdding = true;
            ShowModal = true;
            IsEdit = false;
        }

        private void HideModal()
        {
            CurrentBank = null;
            IsAdding = false;
            ShowModal = false;
            IsEdit = false;
        }
        private void ShowWarning(string summary, string detail)
        {
            NotificationService.Notify(new NotificationMessage
            {
                Severity = NotificationSeverity.Warning,
                Summary = $"⚠️ {summary}",
                Detail = detail,
                Duration = 3000,
                Style = MsgStyle("#fff3cd", "#856404", "#ffc107")
            });
        }
        private void ShowSuccess(string summary, string detail)
        {
            NotificationService.Notify(new NotificationMessage
            {
                Severity = NotificationSeverity.Success,
                Summary = summary,
                Detail = detail,
                Duration = 3000,
                Style = MsgStyle("#d4edda", "#155724", "#28a745")
            });
        }
        private void ShowError(string summary, string detail)
        {
            NotificationService.Notify(new NotificationMessage
            {
                Severity = NotificationSeverity.Error,
                Summary = summary,
                Detail = detail,
                Duration = 3000,
                Style = MsgStyle("#f8d7da", "#721c24", "#dc3545")
            });
        }
        private string MsgStyle(string bgColor, string textColor, string borderColor)
        {
            return $@"
                    background-color: {bgColor};
                    color: {textColor};
                    font-weight: 600;
                    font-size: 14px;
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                    position: fixed;
                    top: 20px;
                    right: 20px;
                    z-index: 9999;
                    min-width: 360px;
                    max-width: 480px;
                    box-shadow: 0 4px 12px rgba(0,0,0,0.2);
                    border-left: 6px solid {borderColor};
                    border-radius: 6px;
                    padding: 16px 20px;
                    display: flex;
                    align-items: flex-start;
                    gap: 12px;";
        }

        private async Task Refreshlist()
        {
            IsLoading = true;
            try
            {
                var BanksFromDb = await BankListDataAccess.GetBankList();
                BankListfromDB = BanksFromDb.ToList();
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

        [Inject]
        protected Radzen.NotificationService Notification { get; set; }

        private async Task DeleteSelectedBank(ItemClickEventArgs e)
        {
            try
            {
                NotificationServiceManager.ShowWarning("Cannot delete this bank.", "Option Restricted.");
                return;
                StateHasChanged();
            }
            catch (Exception ex)
            {
                //Console.WriteLine($"Error fetching Data: {ex.Message}");
                NotificationServiceManager.ShowError("System Error Message", $"Failed to delete: {ex.Message}.");
            }
        }

    }
}