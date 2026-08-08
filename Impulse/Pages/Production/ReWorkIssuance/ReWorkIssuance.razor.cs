using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Web;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Production.ReWorkIssuance
{
    public partial class ReWorkIssuance : ComponentBase
    {
        [Inject] private IReWorkIssuanceService ReworkService { get; set; } = default!;
        [Inject] private Radzen.NotificationService NotificationService { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = default!;

        // Lookups & Data
        public List<LookupItemString> Articles { get; set; } = new List<LookupItemString>();
        public List<LookupItemInt> RepairTypes { get; set; } = new List<LookupItemInt>();
        public List<LookupItemInt> EligibleMakers { get; set; } = new List<LookupItemInt>();

        public List<ReWorkAvailableItem> AvailableLots { get; set; } = new List<ReWorkAvailableItem>();
        public List<ReWorkStagedItem> StagedLots { get; set; } = new List<ReWorkStagedItem>();

        // Selections
        public LookupItemString? SelectedArticle { get; set; } = null;
        public LookupItemInt? SelectedMaker { get; set; } = null;
        public ReWorkIssuanceFilter Filter { get; set; } = new ReWorkIssuanceFilter();

        // State & Flags
        public bool IsLoadingAvailable { get; set; } = false;
        public bool IsSaving { get; set; } = false;
        public bool PrintPTCAfterSave { get; set; } = false;
        public string LastGeneratedLotNo { get; set; } = string.Empty;
        public string LastMakerIssNo { get; set; } = string.Empty;

        // Computed Totals
        public decimal TotalAvailableQty => AvailableLots.Sum(x => x.BalanceQty);
        public decimal TotalStagedQty => StagedLots.Sum(x => x.Qty);

        protected override async Task OnInitializedAsync()
        {
            try
            {
                Articles = await ReworkService.GetArticlesAsync();
                RepairTypes = await ReworkService.GetRepairTypesAsync();

                Filter.DateRangeIndex = 0;
                Filter.DtFrom = DateTime.Today;
                Filter.DtTo = DateTime.Today;
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

        public async Task<IEnumerable<LookupItemString>> SearchArticles(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return await Task.FromResult(Articles);
            }

            return await Task.FromResult(Articles.Where(a =>
                a.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                a.Id.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        public async Task<IEnumerable<LookupItemInt>> SearchMakers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return await Task.FromResult(EligibleMakers);
            }

            return await Task.FromResult(EligibleMakers.Where(m =>
                m.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        public async Task OnArticleChanged(LookupItemString? article)
        {
            SelectedArticle = article;
            Filter.ArticleId = article?.Id ?? string.Empty;
            ClearStaging();
            await LoadAvailableLotsAsync();
        }

        public async Task OnMakerChanged(LookupItemInt? maker)
        {
            SelectedMaker = maker;
            await Task.CompletedTask;
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
                    // Keep existing DtFrom/DtTo
                    break;
            }

            await LoadAvailableLotsAsync();
        }

        public async Task OnLotNoKeyDown(KeyboardEventArgs e)
        {
            if (e.Key == "Enter")
            {
                await LoadAvailableLotsAsync();
            }
        }

        public async Task LoadAvailableLotsAsync()
        {
            if (SelectedArticle == null || string.IsNullOrWhiteSpace(SelectedArticle.Id))
            {
                AvailableLots.Clear();
                return;
            }

            IsLoadingAvailable = true;
            try
            {
                var list = await ReworkService.GetAvailableReWorkLotsAsync(Filter);

                // Exclude any lots already staged
                var stagedIds = new HashSet<long>(StagedLots.Select(s => s.EntryID));
                AvailableLots = list.Where(x => !stagedIds.Contains(x.EntryID)).ToList();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Failed to Load Lots",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoadingAvailable = false;
            }
        }

        public async Task StageLot(ReWorkAvailableItem item)
        {
            if (StagedLots.Any())
            {
                var first = StagedLots.First();

                if (!string.Equals(item.OrderNo, first.OrderNo, StringComparison.OrdinalIgnoreCase))
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Warning,
                        Summary = "Mismatched Order No",
                        Detail = $"Cannot add Lot #{item.LotNo}. Customer Order '{item.OrderNo}' does not match batch Order '{first.OrderNo}'.",
                        Duration = 5000
                    });
                    return;
                }

                if (item.VendID != first.VendID)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Warning,
                        Summary = "Mismatched Maker",
                        Detail = $"Cannot add Lot #{item.LotNo}. Origin Maker does not match batch Origin Maker.",
                        Duration = 5000
                    });
                    return;
                }

                if (item.Repair_RefID != first.Repair_RefID)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Warning,
                        Summary = "Mismatched Repair Type",
                        Detail = $"Cannot add Lot #{item.LotNo}. Repair Type '{item.RepairType}' does not match batch Repair Type '{first.RepairType}'.",
                        Duration = 5000
                    });
                    return;
                }
            }

            StagedLots.Add(new ReWorkStagedItem
            {
                EntryID = item.EntryID,
                LotNo = item.LotNo,
                OrderNo = item.OrderNo,
                VendID = item.VendID,
                DisplayMaker = item.DisplayMaker,
                Repair_RefID = item.Repair_RefID,
                RepairType = item.RepairType,
                Qty = item.BalanceQty,
                Rate = item.Rate
            });

            AvailableLots.Remove(item);

            if (StagedLots.Count == 1)
            {
                try
                {
                    EligibleMakers = await ReworkService.GetMakersForRepairProcessAsync(item.Repair_RefID);
                    SelectedMaker = null;
                }
                catch (Exception ex)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Error,
                        Summary = "Failed to Load Makers",
                        Detail = ex.Message,
                        Duration = 4000
                    });
                }
            }
        }

        public async Task UnstageLot(ReWorkStagedItem item)
        {
            StagedLots.Remove(item);
            if (!StagedLots.Any())
            {
                EligibleMakers.Clear();
                SelectedMaker = null;
            }

            await LoadAvailableLotsAsync();
        }

        public void ClearStaging()
        {
            StagedLots.Clear();
            EligibleMakers.Clear();
            SelectedMaker = null;
        }

        public async Task PostReWorkIssuance()
        {
            if (!StagedLots.Any())
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "No Lots Staged",
                    Detail = "Please stage at least one rework lot before issuing.",
                    Duration = 4000
                });
                return;
            }

            if (SelectedMaker == null)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "No Maker Selected",
                    Detail = "Please select a maker for this rework issuance.",
                    Duration = 4000
                });
                return;
            }

            if (SelectedArticle == null)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "No Article Selected",
                    Detail = "Please select an article.",
                    Duration = 4000
                });
                return;
            }

            IsSaving = true;
            try
            {
                string userName = await GetCurrentUserName();
                string machineName = Environment.MachineName;

                var req = new SaveReWorkIssuanceRequest
                {
                    ItemID = SelectedArticle.Id,
                    MakerID = SelectedMaker.Id,
                    StagedItems = StagedLots,
                    PrintPTC = PrintPTCAfterSave,
                    UserName = userName,
                    MachineName = machineName
                };

                var result = await ReworkService.SaveReWorkIssuanceAsync(req);

                if (result.Success)
                {
                    LastGeneratedLotNo = result.NewLotNo;
                    LastMakerIssNo = result.MakerIssNo;

                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Success,
                        Summary = "Rework Issued Successfully",
                        Detail = $"Generated Lot #{result.NewLotNo} (Issuance: {result.MakerIssNo}).",
                        Duration = 6000
                    });

                    if (PrintPTCAfterSave)
                    {
                        await PrintPTCForLot(result.NewLotNo);
                    }

                    ClearStaging();
                    await LoadAvailableLotsAsync();
                }
                else
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Error,
                        Summary = "Issuance Failed",
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
                    Summary = "Operation Error",
                    Detail = ex.Message,
                    Duration = 6000
                });
            }
            finally
            {
                IsSaving = false;
            }
        }

        public async Task PrintPTCForLot(string lotNo)
        {
            if (string.IsNullOrWhiteSpace(lotNo)) return;

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "PTCQel.rpt",
                SelectionFormula = $"{{VFOrderItemPTC.LotNo}}='{lotNo}'"
            });
        }

        public async Task PrintPendingReport()
        {
            string selection = "{@Balance}>0";
            if (SelectedArticle != null && !string.IsNullOrWhiteSpace(SelectedArticle.Id))
            {
                selection += $" AND {{VVendRcvdDetailReWorkDetail.ItemCode}}='{SelectedArticle.Id}'";
            }

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "PendingReWorks.rpt",
                SelectionFormula = selection
            });
        }

        private async Task<string> GetCurrentUserName()
        {
            try
            {
                var authState = await AuthStateProvider.GetAuthenticationStateAsync();
                var user = authState.User;
                return user.Identity?.IsAuthenticated == true
                    ? user.Identity.Name ?? "System"
                    : "System";
            }
            catch
            {
                return "System";
            }
        }
    }
}
