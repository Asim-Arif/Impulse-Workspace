using Azure.Identity;
using BlazorBootstrap;
using BlazorContextMenu;
using CIP.Pages;
using DataAccessLibrary;
using DataAccessLibrary.DAC.Accounts;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.EntityFrameworkCore.Infrastructure.Internal;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Microsoft.JSInterop;
using Microsoft.VisualBasic;
using Radzen;
using Radzen.Blazor.Markdown;
using System.Collections.ObjectModel;
using System.Security.AccessControl;
using System.Threading.Tasks;
using OpenXmlMenu = DocumentFormat.OpenXml.Office2010.CustomUI;
using CIP.Services;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CIP.Pages.Accounts
{
    public partial class TransactionRegister_Adjustment : ComponentBase
    {
        private bool IsLoading=true;
        [Inject]
        private IVouchersDataAccess IVouchersDataAccess { get; set; }
        [Inject]
        private IDBHelper IDBHelper { get; set; }
        [Inject]
        protected IAccountReportingAccess AccountReportingAccess { get; set; }
        [Inject]
        protected IFinancialStatementService _financialstatementService { get; set; }
        [Inject]
        protected IReportNavigationService ReportNavigationService { get; set; }

        private List<GenericDropDownModel> Accounts = new List<GenericDropDownModel>();
        private List<AccountsReportingModel> AccountsList = new List<AccountsReportingModel>();
        private AccountsReportingModel CurrentAccount = new AccountsReportingModel();
        private Vouchers_Adjustment_Transaction_Model CurrentVoucher = new Vouchers_Adjustment_Transaction_Model();
        private List<Vouchers_Adjustment_Transaction_Model> LedgerDatafromDB { get; set; } = new List<Vouchers_Adjustment_Transaction_Model>();
        private List<AccountsReportingModel> LedgerDatafromDB_temp { get; set; } = new List<AccountsReportingModel>();
        private DateTime DtFrom { get; set; } = DateTime.Now;
        private DateTime DtTo { get; set; } = DateTime.Now;

        private string? StrVoucherType = null;

        private bool IsJournal = true;
        private string userName;
        [Parameter] public string? p_VchrNo { get; set; }
        private string? pVoucherNo;
        protected override async Task OnInitializedAsync()
        {
            IsLoading = true;
            try
            {
                IsLoading = true;
                Accounts = await IVouchersDataAccess.GetValuesForDropDown("Accounts", "AccNo", "AccTitle", " WHERE Parent=0 ORDER BY AccNo");
                string maxDateString = DateTime.Today.ToString(); //await IDBHelper.getSingleStringValue("MAX(VDate)", "Vouchers");
                CurrentAccount.DTFrom = DateTime.Parse(maxDateString); //new DateTime(DateTime.Now.Year, 7, 1);
                CurrentAccount.DTTo = DateTime.Parse(maxDateString);

                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                var user = authState.User;
                userName = user.Identity.Name;

                AccountsList = Accounts
                    .Select(a => new AccountsReportingModel
                    {
                        AccNo = a.DropDownValue_ID,
                        AccTitle = a.DropDownValue_Description 
                    })
                    .ToList();
                IsLoading = false;
                //Console.WriteLine($"Voucher parameter: {p_VchrNo}");
                if (!string.IsNullOrEmpty(p_VchrNo))
                {
                    IsLoading = true;
                    pVoucherNo = p_VchrNo;

                    string VchrDT = await IDBHelper.getSingleStringValue("MAX(VDate)", "Vouchers", $"WHERE VchrNo='{@pVoucherNo}'");

                    CurrentAccount.DTFrom = DateTime.Parse(VchrDT);
                    CurrentAccount.DTTo = DateTime.Parse(VchrDT);

                    RefreshLedger();
                    await JS.InvokeVoidAsync("scrollToVoucher", p_VchrNo);
                }

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

        private IEnumerable<IGrouping<string, Vouchers_Adjustment_Transaction_Model>> LedgerGroups;
        public string VoucherType { get; set; }
        private async Task RefreshLedger()
        {
            IsLoading = true;
            try
            {
                var strcond ="";

                strcond = " AND ( ";
                if (IsJournal ==true) 
                {
                    strcond = strcond + " LEFT(VchrNo,2)='JV' OR";
                }
                strcond = strcond.Substring(0, strcond.Length - 3);
                strcond = strcond + " ";

                LedgerDatafromDB.Clear();
                LedgerGroups = null;

                var LedgerDataFromDb = await _financialstatementService.GetTransactionData_Adjustment(CurrentAccount.DTFrom,CurrentAccount.DTTo);
                LedgerDatafromDB = LedgerDataFromDb.ToList();

                StateHasChanged();

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
        private AccountsReportingModel SelectedVoucher;
        private bool ShowModal = false;
        private bool IsEdit = false;
        string lastVoucher = null;
        private async Task EditSelectedVoucher(ItemClickEventArgs e)
        {
            var account = e.Data as AccountsReportingModel;
            SelectedVoucher = account;
            var LedgerDataFromDb = await AccountReportingAccess.GetVoucherData(SelectedVoucher.VchrNo);
            LedgerDatafromDB_temp = LedgerDataFromDb.ToList();

            if (SelectedVoucher != null)
            {
                CurrentAccount = new AccountsReportingModel
                {
                    AccNo = SelectedVoucher.AccNo,
                    AccTitle = SelectedVoucher.AccTitle,
                    Debit = SelectedVoucher.Debit,
                    Credit = SelectedVoucher.Credit,
                    VchrNo = SelectedVoucher.VchrNo,
                    VDate = SelectedVoucher.VDate,
                    Description = SelectedVoucher.Description,
                    GeneratedBy = SelectedVoucher.GeneratedBy,
                    MachineName = SelectedVoucher.MachineName,
                    UserName = userName,
                };

                ShowModal = true;
                IsEdit = true;
            }
        }

        private async Task DeleteSelectedVoucher(ItemClickEventArgs e)
        {
            var account = e.Data as Vouchers_Adjustment_Transaction_Model;
            CurrentVoucher = account;

            if (CurrentVoucher != null)
            {
                var LedgerDataFromDb = await _financialstatementService.DeleteVoucher_Adjustment(CurrentVoucher.VchrNo);
                RefreshLedger();
                StateHasChanged();
            }
        }

        private async Task DeleteVoucher()
        {
            if (IsEdit==false)
                return;
            try
            {
                if (string.IsNullOrWhiteSpace(SelectedVoucher.DeleteReason))
                {
                    await JS.InvokeVoidAsync("alert", "Please Enter Reason of Deletion.");
                    return;
                }

                await AccountReportingAccess.DeleteVoucher(SelectedVoucher, CurrentVoucher.VchrNo,false);

                ShowModal = false;
                //SelectedVoucher = null;
                CurrentVoucher = null;
                HideModal();
                RefreshLedger();
                StateHasChanged();
            }
            catch (Exception ex)
            {
                ShowModal =false;
            }
        }

        private void HideModal()
        {
            SelectedVoucher = null;
            ShowModal = false;
            IsEdit = false;
        }

        private async Task PrintVoucher(ItemClickEventArgs e)
        {
            var vchr = e.Data as Vouchers_Adjustment_Transaction_Model;
            CurrentVoucher = vchr;

            if (CurrentVoucher != null)
            {
                string cond = "{Electricity_Invoices_Detail.VchrNo}='" + CurrentVoucher.VchrNo + "'";
                //ReportNavigationService.PrintReportAsync("JV_Voucher_Adjustment.rpt", cond);
                ReportNavigationService.PrintReportAsync_Old("JV_Voucher_Adjustment.rpt", cond);
            }
        }

        private void PrintTransaction()
        {
            if (CurrentAccount != null)
            {
                string cond = "{VLedger.DT}=#" + CurrentAccount.DTFrom + "# TO #" + CurrentAccount.DTTo + "#";
                //ReportNavigationService.PrintReportAsync("VoucherList_Adjustment.rpt", cond);
                ReportNavigationService.PrintReportAsync_Old("VoucherList_Adjustment.rpt", cond);
            }
        }

    }
}