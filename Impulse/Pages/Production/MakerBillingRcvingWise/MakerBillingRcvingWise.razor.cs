using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Production.MakerBillingRcvingWise
{
    public partial class MakerBillingRcvingWise : ComponentBase
    {
        [Inject] private IMakerBillingService BillingService { get; set; } = default!;
        [Inject] private Radzen.NotificationService NotificationService { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigationService { get; set; } = default!;

        // Lookups & Filters
        public MakerBillingLookups Lookups { get; set; } = new MakerBillingLookups();
        public LookupItemInt? SelectedMaker { get; set; } = null;
        public LookupItemString? SelectedDebitHead { get; set; } = null;
        public MakerBillingFilter Filter { get; set; } = new MakerBillingFilter();

        // 6 Tab Datasets
        public List<MakerBillingLotItem> Lots { get; set; } = new List<MakerBillingLotItem>();
        public List<MakerRepairItem> RepairItems { get; set; } = new List<MakerRepairItem>();
        public List<MakerMaterialDedItem> MaterialDedItems { get; set; } = new List<MakerMaterialDedItem>();
        public List<MakerManualDedItem> ManualDedItems { get; set; } = new List<MakerManualDedItem>();
        public List<MakerHiddenLotItem> HiddenLots { get; set; } = new List<MakerHiddenLotItem>();
        public List<MakerHiddenRMItem> HiddenRMItems { get; set; } = new List<MakerHiddenRMItem>();

        // Deduction Balances
        public decimal ShortTermDeduction { get; set; } = 0;
        public decimal LongTermDeduction { get; set; } = 0;

        // UI State
        public int ActiveTab { get; set; } = 0;
        public bool IsLoading { get; set; } = false;
        public bool IsPosting { get; set; } = false;

        // Computed Properties
        public decimal CalculatedRepairAmount => RepairItems.Where(x => x.IsSelected).Sum(x => x.WastageAmt);
        public decimal CalculatedMaterialDeductionAmount => MaterialDedItems.Where(x => x.IsSelected).Sum(x => x.Amount);
        public decimal CalculatedManualDeductionAmount => ManualDedItems.Sum(x => x.Amount);

        protected override async Task OnInitializedAsync()
        {
            try
            {
                Lookups = await BillingService.GetLookupsAsync();

                Filter.DateRangeIndex = 0;
                Filter.DtFrom = DateTime.Today;
                Filter.DtTo = DateTime.Today;
                Filter.PostingDate = DateTime.Today;

                if (Lookups.DebitHeads.Any())
                {
                    SelectedDebitHead = Lookups.DebitHeads.First();
                    Filter.DebitHeadAccNo = SelectedDebitHead.Id;
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Initialization Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public async Task<IEnumerable<LookupItemInt>> SearchMakers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return await Task.FromResult(Lookups.Makers);
            }

            return await Task.FromResult(Lookups.Makers.Where(m =>
                m.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        public async Task<IEnumerable<LookupItemString>> SearchDebitHeads(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return await Task.FromResult(Lookups.DebitHeads);
            }

            return await Task.FromResult(Lookups.DebitHeads.Where(h =>
                h.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                h.Id.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        public async Task OnMakerChanged(LookupItemInt? maker)
        {
            SelectedMaker = maker;
            Filter.VendID = maker?.Id ?? 0;
            await LoadAllDataAsync();
        }

        public void OnDebitHeadChanged(LookupItemString? head)
        {
            SelectedDebitHead = head;
            Filter.DebitHeadAccNo = head?.Id ?? string.Empty;
        }

        public async Task OnDateRangeChanged()
        {
            switch (Filter.DateRangeIndex)
            {
                case 0:
                    Filter.DtFrom = DateTime.Today;
                    Filter.DtTo = DateTime.Today;
                    break;
                case 1:
                    Filter.DtFrom = DateTime.Today.AddDays(-7);
                    Filter.DtTo = DateTime.Today;
                    break;
                case 2:
                    Filter.DtFrom = DateTime.Today.AddDays(-15);
                    Filter.DtTo = DateTime.Today;
                    break;
                case 3:
                    Filter.DtFrom = DateTime.Today.AddDays(-30);
                    Filter.DtTo = DateTime.Today;
                    break;
                case 4:
                    Filter.DtFrom = DateTime.Today.AddDays(-60);
                    Filter.DtTo = DateTime.Today;
                    break;
                case 5:
                    Filter.DtFrom = DateTime.Today.AddDays(-90);
                    Filter.DtTo = DateTime.Today;
                    break;
                case 6:
                    Filter.DtFrom = DateTime.Today.AddYears(-1);
                    Filter.DtTo = DateTime.Today;
                    break;
                case 7:
                    Filter.DtFrom = DateTime.Today.AddYears(-5);
                    Filter.DtTo = DateTime.Today;
                    break;
                case 8:
                    // Custom range
                    break;
            }

            await LoadAllDataAsync();
        }

        public async Task LoadAllDataAsync()
        {
            if (Filter.VendID <= 0)
            {
                Lots.Clear();
                RepairItems.Clear();
                MaterialDedItems.Clear();
                HiddenLots.Clear();
                HiddenRMItems.Clear();
                ShortTermDeduction = 0;
                LongTermDeduction = 0;
                return;
            }

            IsLoading = true;
            try
            {
                Lots = await BillingService.GetMakerBillingLotsAsync(Filter);
                RepairItems = await BillingService.GetRepairWastageItemsAsync(Filter);
                MaterialDedItems = await BillingService.GetMaterialDeductionsAsync(Filter.VendID, Filter.MasterPONo);
                HiddenLots = await BillingService.GetHiddenLotsAsync(Filter.VendID);
                HiddenRMItems = await BillingService.GetHiddenRMItemsAsync(Filter.VendID);

                ShortTermDeduction = await BillingService.GetShortTermDeductionAsync(Filter.VendID);
                LongTermDeduction = await BillingService.GetLongTermBalanceAsync(Filter.VendID);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Failed to Load Maker Data",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        public void SetTab(int tabIndex)
        {
            ActiveTab = tabIndex;
        }

        public async Task HandleHideLot(MakerBillingLotItem lot)
        {
            try
            {
                await BillingService.HideLotFromBillingAsync(lot.EntryID);
                Lots.Remove(lot);

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Info,
                    Summary = "Lot Hidden",
                    Detail = $"Lot #{lot.LotNo} moved to Hidden Lots tab.",
                    Duration = 3000
                });

                HiddenLots = await BillingService.GetHiddenLotsAsync(Filter.VendID);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Operation Failed",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public async Task HandleHideRM(MakerMaterialDedItem rm)
        {
            try
            {
                await BillingService.HideRMFromBillingAsync(rm.EntryID);
                MaterialDedItems.Remove(rm);

                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Info,
                    Summary = "Material Hidden",
                    Detail = $"Material #{rm.IssNo} moved to RM Hidden tab.",
                    Duration = 3000
                });

                HiddenRMItems = await BillingService.GetHiddenRMItemsAsync(Filter.VendID);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Operation Failed",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public async Task HandlePrintPTC(MakerBillingLotItem lot)
        {
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "VendReceivings.rpt",
                SelectionFormula = $"{{VendRcvd.EntryID}} = {lot.EntryID}"
            });
        }

        public async Task HandlePrintPTCFromHidden(MakerHiddenLotItem hidden)
        {
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "VendReceivings.rpt",
                SelectionFormula = $"{{VendRcvd.EntryID}} = {hidden.VRD_RefID}"
            });
        }

        public async Task PrintMaterialDeductionsReport()
        {
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "RawMaterialIssuanceToMaker.rpt",
                SelectionFormula = $"{{RawMaterialIssuance.VendID}} = {Filter.VendID}"
            });
        }

        public async Task PrintUnPostedBills()
        {
            if (Filter.VendID <= 0)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Please Select Maker",
                    Detail = "Select a maker first to print un-posted bills.",
                    Duration = 3000
                });
                return;
            }

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "MakerUnPostedBills.rpt",
                SelectionFormula = string.Empty,
                Parameters = new Dictionary<string, object>
                {
                    { "@MakerID", Filter.VendID },
                    { "@DTFrom", Filter.DtFrom },
                    { "@DTTo", Filter.DtTo },
                    { "@CatID", string.IsNullOrWhiteSpace(Filter.CatIDs) ? "0" : Filter.CatIDs },
                    { "@ProcessID", 0 },
                    { "@MasterPONo", Filter.MasterPONo ?? string.Empty }
                }
            });
        }

        public async Task PrintAllMakersSummary()
        {
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "MakerBillingSummary.rpt",
                SelectionFormula = string.Empty
            });
        }

        public async Task PostMakerBill()
        {
            var selectedLots = Lots.Where(x => x.IsSelected).ToList();
            if (!selectedLots.Any())
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "No Lots Selected",
                    Detail = "Please select at least one lot to post in the bill.",
                    Duration = 4000
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(Filter.DebitHeadAccNo))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Missing Debit Head",
                    Detail = "Please select a Debit Head account before posting.",
                    Duration = 4000
                });
                return;
            }

            IsPosting = true;
            try
            {
                var req = new PostMakerBillRequest
                {
                    VendID = Filter.VendID,
                    DebitHeadAccNo = Filter.DebitHeadAccNo,
                    PostingDate = Filter.PostingDate,
                    DtFrom = Filter.DtFrom,
                    DtTo = Filter.DtTo,
                    RejectionRelaxation = Filter.RejectionRelaxation,
                    ShortTermDeduction = ShortTermDeduction,
                    LongTermDeduction = LongTermDeduction,
                    RepairAmount = CalculatedRepairAmount,
                    MaterialDeductionAmount = CalculatedMaterialDeductionAmount,
                    ManualDeductionAmount = CalculatedManualDeductionAmount,
                    SelectedLots = selectedLots,
                    SelectedMaterialDeductions = MaterialDedItems.Where(x => x.IsSelected).ToList(),
                    ManualDeductions = ManualDedItems
                };

                var result = await BillingService.PostMakerBillAsync(req);

                if (result.Success)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Success,
                        Summary = "Bill Posted Successfully",
                        Detail = result.Message,
                        Duration = 6000
                    });

                    // Print Posted Bill Report
                    await ReportNavigationService.PrintReportAsync(new ReportRequest
                    {
                        ReportName = "MakerPostedBill.rpt",
                        SelectionFormula = $"{{MakerPostedBills.EntryID}} = {result.BillEntryID}"
                    });

                    await LoadAllDataAsync();
                }
                else
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Error,
                        Summary = "Posting Failed",
                        Detail = result.Message,
                        Duration = 6000
                    });
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Posting Error",
                    Detail = ex.Message,
                    Duration = 6000
                });
            }
            finally
            {
                IsPosting = false;
            }
        }
    }
}
