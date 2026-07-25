using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.AspNetCore.Components;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Stock.VendorBilling
{
    public partial class VendorBilling : ComponentBase
    {
        // -----------------------------------------------------------------------
        // Injected Services
        // -----------------------------------------------------------------------
        [Inject] private Services.Stock.IVendorBillingService VendorBillingService { get; set; }
        [Inject] private Services.IVoucherService VoucherService { get; set; }
        [Inject] private Services.IAuditService AuditService { get; set; }
        [Inject] private Services.INotificationService NotificationService { get; set; }
        [Inject] private Services.IReportNavigationService ReportNavigation { get; set; }

        // -----------------------------------------------------------------------
        // State
        // -----------------------------------------------------------------------
        private VendorBillingFilterModel Filter { get; set; } = new();
        private VendorBillingSummaryModel Summary { get; set; } = new();

        private List<VendorBillingRowViewModel> BillingRows { get; set; } = new();
        private List<MaterialDeductionRowViewModel> MaterialRows { get; set; } = new();
        private List<ManualDeductionRowViewModel> ManualRows { get; set; } = new();

        private List<VendorLookupModel> AllVendors { get; set; } = new();
        private List<AccountLookupModel> AllAccounts { get; set; } = new();

        private VendorLookupModel? SelectedVendor { get; set; }
        private AccountLookupModel? SelectedDebitAccount { get; set; }

        private int ActiveTab { get; set; } = 0;
        private bool IsLoadingBilling { get; set; } = false;
        private bool IsLoadingMaterial { get; set; } = false;
        private bool IsPosting { get; set; } = false;

        private decimal NetAmountToPost => (decimal)BillingRows.Where(r => r.IsChecked).Sum(r => r.Value)
            - Summary.ShortTermDeduction
            - Summary.LongTermDeduction
            - (decimal)MaterialRows.Sum(r => r.Amount)
            - ManualRows.Sum(r => r.Amount);

        // -----------------------------------------------------------------------
        // Lifecycle
        // -----------------------------------------------------------------------
        protected override async Task OnInitializedAsync()
        {
            AllVendors = (await VendorBillingService.GetVendorsAsync()).ToList();
            AllAccounts = (await VendorBillingService.GetActiveAccountsAsync()).ToList();
            OnDateRangeChanged(); // set default dates for DateRangeIndex = 1 (Last 15 days)
        }

        // -----------------------------------------------------------------------
        // Lookups / Typeahead
        // -----------------------------------------------------------------------
        private Task<IEnumerable<VendorLookupModel>> SearchVendors(string searchText)
        {
            var results = string.IsNullOrWhiteSpace(searchText)
                ? AllVendors
                : AllVendors.Where(v => v.AccTitle.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList();
            return Task.FromResult<IEnumerable<VendorLookupModel>>(results);
        }

        private Task<IEnumerable<AccountLookupModel>> SearchAccounts(string searchText)
        {
            var results = string.IsNullOrWhiteSpace(searchText)
                ? AllAccounts
                : AllAccounts.Where(a => a.AccTitle.Contains(searchText, StringComparison.OrdinalIgnoreCase)
                                      || a.AccNo.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList();
            return Task.FromResult<IEnumerable<AccountLookupModel>>(results);
        }

        private async Task OnVendorSelected()
        {
            if (SelectedVendor != null)
            {
                Filter.VendorAccNo = SelectedVendor.AccNo;
                Filter.VendorAccTitle = SelectedVendor.AccTitle;
                await LoadDataAsync();
            }
            else
            {
                Filter.VendorAccNo = string.Empty;
                BillingRows.Clear();
                MaterialRows.Clear();
                Summary = new();
            }
        }

        private void OnDebitHeadSelected()
        {
            if (SelectedDebitAccount != null)
            {
                Filter.DebitHeadAccNo = SelectedDebitAccount.AccNo;
                Filter.DebitHeadAccTitle = SelectedDebitAccount.AccTitle;
            }
            else
            {
                Filter.DebitHeadAccNo = string.Empty;
            }
        }

        // -----------------------------------------------------------------------
        // Date Range
        // -----------------------------------------------------------------------
        private void OnDateRangeChanged()
        {
            Filter.DateTo = DateTime.Today;
            Filter.DateFrom = Filter.DateRangeIndex switch
            {
                0 => DateTime.Today,
                1 => DateTime.Today.AddDays(-15),
                2 => DateTime.Today.AddDays(-30),
                3 => DateTime.Today.AddDays(-60),
                4 => DateTime.Today.AddDays(-90),
                _ => Filter.DateFrom // custom – keep existing
            };
        }

        // -----------------------------------------------------------------------
        // Data Loading
        // -----------------------------------------------------------------------
        private async Task LoadDataAsync()
        {
            if (string.IsNullOrEmpty(Filter.VendorAccNo)) return;

            try
            {
                IsLoadingBilling = true;
                IsLoadingMaterial = true;
                BillingRows.Clear();
                MaterialRows.Clear();
                StateHasChanged();

                var billingTask = VendorBillingService.GetUnpostedReceivingsAsync(Filter.VendorAccNo, Filter.DateFrom, Filter.DateTo);
                var materialTask = VendorBillingService.GetMaterialIssuancesAsync(Filter.VendorAccNo);
                var shortTermTask = VendorBillingService.GetShortTermDeductionAsync(Filter.VendorAccNo);
                var longTermTask = VendorBillingService.GetLongTermDeductionAsync(Filter.VendorAccNo);

                await Task.WhenAll(billingTask, materialTask, shortTermTask, longTermTask);

                BillingRows = billingTask.Result.ToList();
                MaterialRows = materialTask.Result.ToList();
                Summary.ShortTermDeduction = shortTermTask.Result;
                Summary.LongTermDeduction = longTermTask.Result;
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", $"Failed to load data: {ex.Message}");
            }
            finally
            {
                IsLoadingBilling = false;
                IsLoadingMaterial = false;
                StateHasChanged();
            }
        }

        // -----------------------------------------------------------------------
        // Tab 1 callbacks
        // -----------------------------------------------------------------------
        private void OnBillingDataChanged() => StateHasChanged();

        private async Task OnRateSaved((int entryId, float newRate) args)
        {
            try
            {
                await VendorBillingService.UpdateReceivingRateAsync(args.entryId, args.newRate);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", $"Failed to save rate: {ex.Message}");
            }
        }

        private async Task OnHideEntry(int entryId)
        {
            try
            {
                string userName = AuditService.GetCurrentUserName();
                string machineName = AuditService.GetClientIpAddress();
                await VendorBillingService.HideFromBillingAsync(entryId, userName, machineName);
                BillingRows.RemoveAll(r => r.EntryID == entryId);
                NotificationService.ShowSuccess("Hidden", "Entry hidden from billing.");
                StateHasChanged();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", $"Failed to hide entry: {ex.Message}");
            }
        }

        // -----------------------------------------------------------------------
        // Tab 3 callback
        // -----------------------------------------------------------------------
        private void OnManualDataChanged() => StateHasChanged();

        // -----------------------------------------------------------------------
        // Print Un-Posted
        // -----------------------------------------------------------------------
        private async Task PrintUnPosted()
        {
            var selectedIds = BillingRows.Where(r => r.IsChecked).Select(r => r.EntryID).ToList();
            if (!selectedIds.Any())
            {
                NotificationService.ShowWarning("No Selection", "Please select at least one row to print.");
                return;
            }
            string idList = string.Join(",", selectedIds);
            await ReportNavigation.PrintReportAsync(new Services.ReportRequest
            {
                ReportName = "VenderUnPostedBill.rpt",
                SelectionFormula = $"{{VVendRcvdDetailPO.EntryID}} in [{idList}]"
            });
        }

        // -----------------------------------------------------------------------
        // Post Selected – Critical business flow
        // -----------------------------------------------------------------------
        private async Task PostSelected()
        {
            // Validations
            if (string.IsNullOrEmpty(Filter.DebitHeadAccNo))
            {
                NotificationService.ShowWarning("Validation", "Please select a Debit Head account.");
                return;
            }
            var checkedRows = BillingRows.Where(r => r.IsChecked).ToList();
            if (!checkedRows.Any())
            {
                NotificationService.ShowWarning("Validation", "No orders selected. Please check at least one row.");
                return;
            }

            decimal selectedValue = (decimal)checkedRows.Sum(r => r.Value);
            decimal materialDed = (decimal)MaterialRows.Sum(r => r.Amount);
            decimal manualDed = ManualRows.Sum(r => r.Amount);
            decimal amountToPost = Math.Round(selectedValue) - Summary.ShortTermDeduction - Summary.LongTermDeduction - materialDed - manualDed;

            if (amountToPost <= 0)
            {
                NotificationService.ShowError("Invalid Amount", $"Net amount to post ({amountToPost:N0}) must be greater than zero.");
                return;
            }

            IsPosting = true;
            StateHasChanged();

            try
            {
                string userName = AuditService.GetCurrentUserName();
                string machineName = AuditService.GetClientIpAddress();

                // Step 1: Get next bill number
                int billNo = await VendorBillingService.GetNextBillNoAsync(Filter.VendorAccNo);

                // Step 2: Build description
                string description = $"Computer Bill No. {Filter.VendorAccTitle} / {billNo} - {Filter.PostingDate.Year}";

                // Step 3: Post Journal Voucher via centralised VoucherService
                var voucher = new VoucherViewModel
                {
                    DT = Filter.PostingDate,
                    Notes = description,
                    JVType = 0,
                    PostedThroughJVForm = false
                };
                // VchrNo carries the voucher type prefix ("JV"); InsertVoucherHead_SP
                // re-generates the final sequence under lock and returns it.
                voucher.VchrNo = await VoucherService.GetNextVchrNo(Filter.PostingDate, "JV");
                // Debit: Debit Head account
                voucher.LineItems.Add(new VoucherLineItemViewModel
                {
                    AccNo = Filter.DebitHeadAccNo,
                    AccTitle = Filter.DebitHeadAccTitle,
                    Description = description,
                    Debit = amountToPost,
                    Credit = 0
                });
                // Credit: Vendor account
                voucher.LineItems.Add(new VoucherLineItemViewModel
                {
                    AccNo = Filter.VendorAccNo,
                    AccTitle = Filter.VendorAccTitle,
                    Description = description,
                    Debit = 0,
                    Credit = amountToPost
                });

                string vchrNo = await VoucherService.SaveVoucherAsync(voucher);

                // Step 4: Save all bill detail records in one DB transaction
                var postModel = new VendorBillingPostModel
                {
                    VendorAccNo = Filter.VendorAccNo,
                    DebitHeadAccNo = Filter.DebitHeadAccNo,
                    PostingDate = Filter.PostingDate,
                    DateFrom = Filter.DateFrom,
                    DateTo = Filter.DateTo,
                    BillNo = billNo,
                    VchrNo = vchrNo,
                    ManualDedAmt = manualDed,
                    STDeduction = Summary.ShortTermDeduction,
                    LTDeduction = Summary.LongTermDeduction,
                    MaterialDedAmt = materialDed,
                    UserName = userName,
                    MachineName = machineName,
                    CheckedEntryIDs = checkedRows.Select(r => r.EntryID).ToList(),
                    MaterialIssuanceEntryIDs = MaterialRows.Select(r => r.EntryID).ToList(),
                    ManualDeductions = ManualRows.ToList()
                };

                await VendorBillingService.SavePostedBillAsync(postModel);

                // Step 5: Notify success + print
                NotificationService.ShowSuccess("Posted Successfully",
                    $"Vendor Bill #{billNo} posted. Voucher: {vchrNo}");

                // Step 6: Print the posted bill report
                await ReportNavigation.PrintReportAsync(new Services.ReportRequest
                {
                    ReportName = "VenderPostedBill.rpt",
                    SelectionFormula = $"{{VenderPostedBills.VchrNo}} = '{vchrNo}'"
                });

                // Step 7: Clear manual deductions + reload
                ManualRows.Clear();
                await LoadDataAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Posting Failed", ex.Message);
            }
            finally
            {
                IsPosting = false;
                StateHasChanged();
            }
        }
    }
}
