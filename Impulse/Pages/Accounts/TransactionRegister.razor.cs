using Azure.Identity;
using BlazorBootstrap;
using BlazorContextMenu;
using DataAccessLibrary;
using DataAccessLibrary.DAC.Accounts;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels;
using Impulse.Services;
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

namespace Impulse.Pages.Accounts
{
    public partial class TransactionRegister : ComponentBase
    {
        private bool IsLoading=true;
        [Inject]
        private IVouchersDataAccess IVouchersDataAccess { get; set; }
        [Inject]
        private IDBHelper IDBHelper { get; set; }
        [Inject]
        protected IAccountReportingAccess AccountReportingAccess { get; set; }
        [Inject]
        private Impulse.Services.IReportNavigationService ReportNavigationService { get; set; } = null!;
        [Inject]
        private Impulse.Services.INotificationService NotificationService { get; set; } = null!;
        private List<GenericDropDownModel> Accounts = new List<GenericDropDownModel>();
        private List<AccountsReportingModel> AccountsList = new List<AccountsReportingModel>();
        private AccountsReportingModel CurrentAccount = new AccountsReportingModel();
        private AccountsReportingModel CurrentVoucher = new AccountsReportingModel();
        private List<AccountsReportingModel> LedgerDatafromDB { get; set; } = new List<AccountsReportingModel>();
        private List<AccountsReportingModel> LedgerDatafromDB_temp { get; set; } = new List<AccountsReportingModel>();
        private DateTime DtFrom { get; set; } = DateTime.Now;
        private DateTime DtTo { get; set; } = DateTime.Now;

        private string? StrVoucherType = null;

        private bool IsJournal = true;
        private bool IsBPV = true;
        private bool IsCPV = true;
        private bool IsCRV = true;
        private bool IsBRV = true;
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
                string maxDateString = await IDBHelper.getSingleStringValue("MAX(VDate)", "Vouchers");

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

        private IEnumerable<IGrouping<string, AccountsReportingModel>> LedgerGroups;
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
                if (IsBPV == true) 
                {
                    strcond = strcond + " LEFT(VchrNo,3)='BPV' OR";
                }
                if (IsCPV == true)
                {
                    strcond = strcond + " LEFT(VchrNo,3)='CPV' OR";
                }
                if (IsCRV == true)
                {
                    strcond = strcond + " LEFT(VchrNo,3)='CRV' OR";
                }
                if (IsBRV == true)
                {
                   strcond = strcond + " LEFT(VchrNo,3)='BRV' OR";
                }

                strcond = strcond.Substring(0, strcond.Length - 3);
                strcond = strcond + " ) ";

                LedgerDatafromDB.Clear();
                LedgerGroups = null;

                var LedgerDataFromDb = await AccountReportingAccess.GetTransactionData(CurrentAccount.DTFrom,CurrentAccount.DTTo, strcond);
                LedgerDatafromDB = LedgerDataFromDb.ToList();
                LedgerGroups = LedgerDatafromDB.GroupBy(x => x.VoucherType);

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
            var account = e.Data as AccountsReportingModel;

            CurrentVoucher = account;

            if (CurrentVoucher != null)
            {
                var LedgerDataFromDb = await AccountReportingAccess.GetVoucherData(CurrentVoucher.VchrNo);
                LedgerDatafromDB_temp = LedgerDataFromDb.ToList();

                SelectedVoucher = new AccountsReportingModel
                {
                    AccNo = account.AccNo,
                    AccTitle = account.AccTitle,
                    Debit = account.Debit,
                    Credit = account.Credit,
                    VchrNo = account.VchrNo,
                    VDate = account.VDate,
                    Description = account.Description,
                    GeneratedBy = account.GeneratedBy,
                    MachineName = account.MachineName,
                    UserName = userName
                };

                ShowModal = true;
                IsEdit = true;
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

        private bool bPrintCheque = false;
        private ChqListModel SelectedCheque = new ChqListModel();

        private async Task PrintVoucher(ItemClickEventArgs e)
        {
            var account = e.Data as AccountsReportingModel;
            if (account != null && !string.IsNullOrEmpty(account.VchrNo))
            {
                try
                {
                    await ReportNavigationService.PrintReportAsync(new ReportRequest
                    {
                        ReportName = "Voucher.rpt",
                        SelectionFormula = $"{{VLedger.VchrNo}}='{account.VchrNo}'"
                    });
                }
                catch (Exception ex)
                {
                    NotificationService.ShowError("Report Error", $"Failed to print voucher: {ex.Message}");
                }
            }
        }

        private async Task ShowPrintChequeModal(ItemClickEventArgs e)
        {
            var account = e.Data as AccountsReportingModel;
            if (account == null || string.IsNullOrEmpty(account.VchrNo)) return;

            SelectedCheque = new ChqListModel
            {
                AccNo = account.AccNo,
                AccTitle = account.AccTitle,
                ChqNo = account.ChqNo,
                Amount = account.Debit > 0 ? account.Debit : account.Credit,
                chqpayee = !string.IsNullOrEmpty(account.Handed_Over_To) ? account.Handed_Over_To : account.AccTitle,
                chqdesignation = string.Empty,
                chqcompany = "IAA",
                chqprintdate = account.ChqDate != DateTime.MinValue ? account.ChqDate : (account.VDate != DateTime.MinValue ? account.VDate : DateTime.Today),
                DeleteDescription = account.VchrNo
            };

            try
            {
                string csnoStr = await IDBHelper.getSingleStringValue("SNo", "VChqLedger", $"WHERE VchrNo='{account.VchrNo}'");
                if (double.TryParse(csnoStr, out double csno))
                {
                    SelectedCheque.CSNo = csno;
                }
            }
            catch { }

            bPrintCheque = true;
            StateHasChanged();
        }

        private void HidePrintCheque()
        {
            bPrintCheque = false;
            StateHasChanged();
        }

        private async Task PrintChequeConfirm()
        {
            try
            {
                string rptName = "";
                if (SelectedCheque.CSNo > 0)
                {
                    rptName = await IDBHelper.getSingleStringValue("ChqFormat", "VChqLedger", $"WHERE SNo={SelectedCheque.CSNo}");
                }

                if (string.IsNullOrWhiteSpace(rptName))
                {
                    rptName = "rptChq.rpt";
                }
                else
                {
                    rptName = rptName.Replace(".rpt", "", StringComparison.OrdinalIgnoreCase) + ".rpt";
                }

                string selectionFormula;
                if (SelectedCheque.CSNo > 0)
                {
                    selectionFormula = $"{{VChqLedger.SNo}}={SelectedCheque.CSNo}";
                }
                else
                {
                    selectionFormula = $"{{VLedger.VchrNo}}='{SelectedCheque.DeleteDescription}'";
                }

                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = rptName,
                    SelectionFormula = selectionFormula
                });

                bPrintCheque = false;
                StateHasChanged();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Report Error", $"Failed to print cheque: {ex.Message}");
            }
        }

        private async Task PrintTransactionRegisterReport()
        {
            try
            {
                var types = new List<string>();
                if (IsJournal) { types.Add("'JV'"); }
                if (IsBPV) { types.Add("'BP'"); types.Add("'BPV'"); }
                if (IsCPV) { types.Add("'CP'"); types.Add("'CPV'"); }
                if (IsCRV) { types.Add("'CR'"); types.Add("'CRV'"); }
                if (IsBRV) { types.Add("'BR'"); types.Add("'BRV'"); }

                string typeFilter = types.Any()
                    ? $" AND (left({{VLedger.VchrNo}}, 2) in [{string.Join(", ", types)}] or left({{VLedger.VchrNo}}, 3) in [{string.Join(", ", types)}])"
                    : "";
                string sel = $"{{VLedger.VDate}} in Date({CurrentAccount.DTFrom.Year}, {CurrentAccount.DTFrom.Month}, {CurrentAccount.DTFrom.Day}) to Date({CurrentAccount.DTTo.Year}, {CurrentAccount.DTTo.Month}, {CurrentAccount.DTTo.Day}){typeFilter}";

                var request = new ReportRequest
                {
                    ReportName = "TransRpt1.rpt",
                    SelectionFormula = sel,
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "Company", $"'IAA'" },
                        { "FromTo", $"'From {CurrentAccount.DTFrom:dd-MMM-yyyy} To {CurrentAccount.DTTo:dd-MMM-yyyy}'" }
                    }
                };

                await ReportNavigationService.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Report Error", $"Failed to print register: {ex.Message}");
            }
        }
    }
}