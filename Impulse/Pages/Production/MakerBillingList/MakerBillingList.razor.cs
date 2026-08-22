using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.Production.MakerBillingList
{
    public partial class MakerBillingList : ComponentBase
    {
        [Inject] public IMakerBillingListService BillingListService { get; set; } = default!;
        [Inject] public IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] public Radzen.NotificationService NotificationService { get; set; } = default!;

        public List<MakerBillingListItem> Bills { get; set; } = new();
        public List<LookupItemInt> Makers { get; set; } = new();
        public LookupItemInt SelectedMaker { get; set; } = new() { Id = 0, Name = "<All Makers>" };

        public List<LookupItemString> CreditAccounts { get; set; } = new();
        public LookupItemString? SelectedCreditAccount { get; set; }

        public MakerBillingListFilter Filter { get; set; } = new();
        public MakerBillingListSummary Summary { get; set; } = new();

        public MakerBillingListItem? SelectedBill { get; set; }
        public MakerBillingListItem? ContextRowItem { get; set; }

        public MakerBillingPaymentVoucherRequest? PaymentRequest { get; set; }
        public bool ShowPaymentModal { get; set; } = false;
        public bool IsPostingPayment { get; set; } = false;
        public bool IsLoading { get; set; } = false;

        public List<MakerBillingListItem> FilteredBills
        {
            get
            {
                if (string.IsNullOrWhiteSpace(Filter.SearchText))
                    return Bills;

                string query = Filter.SearchText.Trim();
                return Bills.Where(b =>
                    (!string.IsNullOrEmpty(b.VenderName) && b.VenderName.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(b.VchrNo) && b.VchrNo.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                    (b.BillNo.HasValue && b.BillNo.Value.ToString().Contains(query)) ||
                    (!string.IsNullOrEmpty(b.VendID1) && b.VendID1.Contains(query, StringComparison.OrdinalIgnoreCase))
                ).ToList();
            }
        }

        protected override async Task OnInitializedAsync()
        {
            await LoadMakersAsync();
            await LoadCreditAccountsAsync();
            OnDateRangeChanged();
            await LoadBillsAsync();
        }

        public async Task<IEnumerable<LookupItemInt>> SearchMakers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Makers;

            return Makers.Where(m => m.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase));
        }

        public async Task<IEnumerable<LookupItemString>> SearchCreditAccounts(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return CreditAccounts;

            return CreditAccounts.Where(a => a.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase));
        }

        public async Task LoadMakersAsync()
        {
            try
            {
                Makers = await BillingListService.GetMakersAsync();
                SelectedMaker = Makers.FirstOrDefault(m => m.Id == 0) ?? new LookupItemInt { Id = 0, Name = "<All Makers>" };
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Failed to Load Makers",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public async Task LoadCreditAccountsAsync()
        {
            try
            {
                CreditAccounts = await BillingListService.GetPaymentCreditAccountsAsync();
                SelectedCreditAccount = CreditAccounts.FirstOrDefault();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Failed to Load Accounts",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public async Task OnMakerChanged(LookupItemInt maker)
        {
            SelectedMaker = maker ?? new LookupItemInt { Id = 0, Name = "<All Makers>" };
            Filter.VendID = SelectedMaker.Id;
            await LoadBillsAsync();
        }

        public void OnCreditAccountChanged(LookupItemString acc)
        {
            SelectedCreditAccount = acc;
            if (PaymentRequest != null && acc != null)
            {
                PaymentRequest.CreditAccountNo = acc.Id;
            }
        }

        public void OnDateRangeChanged()
        {
            DateTime today = DateTime.Today;
            switch (Filter.DateRangeIndex)
            {
                case 0: // Today
                    Filter.DtFrom = today;
                    Filter.DtTo = today;
                    break;
                case 1: // Last 15 Days
                    Filter.DtFrom = today.AddDays(-15);
                    Filter.DtTo = today;
                    break;
                case 2: // Last 30 Days
                    Filter.DtFrom = today.AddDays(-30);
                    Filter.DtTo = today;
                    break;
                case 3: // Last 60 Days
                    Filter.DtFrom = today.AddDays(-60);
                    Filter.DtTo = today;
                    break;
                case 4: // Last 90 Days
                    Filter.DtFrom = today.AddDays(-90);
                    Filter.DtTo = today;
                    break;
                case 5: // Custom Range
                    break;
            }
        }

        public async Task LoadBillsAsync()
        {
            IsLoading = true;
            try
            {
                Bills = await BillingListService.GetPostedBillsListAsync(Filter);
                ComputeSummary();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Bills",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void ComputeSummary()
        {
            Summary.TotalBills = Bills.Count;
            Summary.TotalQty = Bills.Sum(b => b.TotalQty);
            Summary.TotalValue = Bills.Sum(b => b.NetPayable);
            Summary.TotalPaid = Bills.Sum(b => b.TotalPaidAmount);
            Summary.TotalBalance = Bills.Sum(b => b.Balance);
        }

        public async Task PrintBillsListReport()
        {
            string formula = $"{{MakerPostedBills.DT}} in Date({Filter.DtFrom:yyyy,MM,dd}) to Date({Filter.DtTo:yyyy,MM,dd})";
            if (Filter.VendID > 0)
            {
                formula += $" and {{MakerPostedBills.VendID}} = {Filter.VendID}";
            }

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "MakerPostedBills_Summary.rpt",
                SelectionFormula = formula,
                Parameters = new Dictionary<string, object>()
            });
        }

        public async Task PrintBill(MakerBillingListItem bill)
        {
            if (bill == null || string.IsNullOrWhiteSpace(bill.VchrNo))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "No Voucher",
                    Detail = "This bill does not have an associated voucher number.",
                    Duration = 3000
                });
                return;
            }

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "MakerPostedBillsRcvWise.rpt",
                SelectionFormula = $"{{MakerPostedBills.VchrNo}} = '{bill.VchrNo}'",
                Parameters = new Dictionary<string, object>()
            });
        }

        public void OpenPaymentVoucherModal(MakerBillingListItem bill)
        {
            if (bill == null) return;

            if (bill.Balance <= 0)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Info,
                    Summary = "Already Paid",
                    Detail = $"Bill #{bill.BillNo} is already fully paid.",
                    Duration = 3000
                });
                return;
            }

            SelectedBill = bill;
            PaymentRequest = new MakerBillingPaymentVoucherRequest
            {
                MPB_RefID = bill.EntryID,
                VendID = bill.VendID,
                MakerName = bill.VenderName,
                BillNo = bill.BillNo,
                BillVoucherNo = bill.VchrNo,
                Amount = bill.Balance,
                PaymentDate = DateTime.Today,
                PaymentMode = "Cash",
                CreditAccountNo = SelectedCreditAccount?.Id ?? string.Empty,
                Remarks = $"Payment against Maker Bill #{bill.BillNo} ({bill.VchrNo})"
            };

            ShowPaymentModal = true;
        }

        public async Task PostPaymentVoucher()
        {
            if (PaymentRequest == null) return;

            if (PaymentRequest.Amount <= 0)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Invalid Amount",
                    Detail = "Payment amount must be greater than zero.",
                    Duration = 3000
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(PaymentRequest.CreditAccountNo))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Select Credit Account",
                    Detail = "Please select a valid Cash in Hand or Bank Credit Account.",
                    Duration = 3000
                });
                return;
            }

            IsPostingPayment = true;
            try
            {
                string voucherNo = await BillingListService.PostPaymentVoucherAsync(PaymentRequest);

                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Payment Voucher Posted",
                    Detail = $"{PaymentRequest.PaymentMode} Voucher #{voucherNo} for Rs. {PaymentRequest.Amount:N2} posted successfully.",
                    Duration = 4000
                });

                ShowPaymentModal = false;
                PaymentRequest = null;
                await LoadBillsAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Payment Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsPostingPayment = false;
            }
        }

        public async Task ContextPrintBill(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                await PrintBill(ContextRowItem);
            }
        }

        public void ContextPaymentVoucher(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                OpenPaymentVoucherModal(ContextRowItem);
            }
        }

        private void ResolveRowItem(ItemClickEventArgs args)
        {
            if (args.Data is MakerBillingListItem item)
            {
                ContextRowItem = item;
            }
            else
            {
                ContextRowItem = SelectedBill;
            }
        }
    }
}
