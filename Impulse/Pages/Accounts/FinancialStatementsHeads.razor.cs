using BlazorBootstrap;
using BlazorContextMenu;
using DataAccessLibrary;
using DataAccessLibrary.DAC.Accounts;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Impulse.Services;
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


namespace Impulse.Pages.Accounts
{
    public partial class FinancialStatementsHeads : ComponentBase, IDisposable
    {
        private FinancialStatementHeadsModel statementmodel = new();

        [Inject]
        private IFinancialStatementService _financialstatementService { get; set; }
        [Inject]
        private INotificationService NotificationServiceManager { get; set; }
        [Inject]
        private IAuditService _auditService { get; set; }

        private List<FinancialStatementHeadsModel> getDBEntries { get; set; } = new List<FinancialStatementHeadsModel>();
        private FinancialStatementHeadsModel CurrentHead =new FinancialStatementHeadsModel();

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

        private void GoToIndexPage()
        {
            Navigation.NavigateTo("/", true);
        }
        private void ShowAddModal()
        {
            CurrentHead = new FinancialStatementHeadsModel { };
            IsAdding = true;
            ShowModal = true;
            IsEdit = false;
        }

        private async Task Refreshlist()
        {
            IsLoading = true;
            try
            {

                string username =_auditService.GetCurrentUserName();
                string machineIP = _auditService.GetClientIpAddress();

                await _financialstatementService.SaveFinancialStatementHeads (username, machineIP);

                var EntriesFromDb = await _financialstatementService.GetFinancialStatementHeadsList();
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

        private async Task ChangeStatus(FinancialStatementHeadsModel item, int itype, bool bMarked)
        {

            CurrentHead = item;
            if (itype == 1)
            {
                CurrentHead.ShowInBS = bMarked;
            }
            else
            {
                CurrentHead.ShowInNTA = bMarked;
            }

            bool bsaved= await _financialstatementService.EditFinancialStatementHeads(CurrentHead, itype, CurrentHead.EntryID);

            if (bsaved) 
            {
                NotificationServiceManager.ShowSuccess("Alert", "Saved.");
            }

            Refreshlist();
            StateHasChanged();
        }

    }
}