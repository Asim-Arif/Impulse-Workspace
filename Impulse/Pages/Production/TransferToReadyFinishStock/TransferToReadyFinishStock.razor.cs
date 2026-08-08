using BlazorContextMenu;
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

namespace Impulse.Pages.Production.TransferToReadyFinishStock
{
    public partial class TransferToReadyFinishStock : ComponentBase
    {
        [Inject] private ITransferToReadyFinishStockService TransferService { get; set; } = default!;
        [Inject] private Radzen.NotificationService NotificationService { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = default!;

        // Lookups & Data
        public List<LookupItemInt> Makers { get; set; } = new List<LookupItemInt>();
        public List<LookupItemString> Articles { get; set; } = new List<LookupItemString>();
        public List<StoreLookupDto> Stores { get; set; } = new List<StoreLookupDto>();
        public List<ShelfLookupDto> Shelves { get; set; } = new List<ShelfLookupDto>();

        public List<ReadyFinishStockAvailableItem> AvailableLots { get; set; } = new List<ReadyFinishStockAvailableItem>();

        // Filters & Selections
        public LookupItemInt? SelectedMaker { get; set; } = null;
        public LookupItemString? SelectedArticle { get; set; } = null;
        public ReadyFinishStockFilter Filter { get; set; } = new ReadyFinishStockFilter();

        // Modal State
        public bool ShowTransferModal { get; set; } = false;
        public int SelectedStoreId { get; set; } = 0;
        public int SelectedShelfId { get; set; } = 0;
        public string TransferRemarks { get; set; } = string.Empty;

        // Flags & Progress
        public bool IsLoading { get; set; } = false;
        public bool IsPosting { get; set; } = false;
        public bool IsAllSelected { get; set; } = false;
        public ReadyFinishStockAvailableItem? ContextRowItem { get; set; } = null;

        // Computed Properties
        public List<ReadyFinishStockAvailableItem> SelectedLots => AvailableLots.Where(x => x.IsSelected).ToList();
        public decimal TotalAvailableQty => AvailableLots.Sum(x => x.AvailQty);
        public decimal TotalSelectedQty => SelectedLots.Sum(x => x.TransferQty);

        protected override async Task OnInitializedAsync()
        {
            try
            {
                Makers = await TransferService.GetMakersAsync();
                Articles = await TransferService.GetArticlesAsync();
                Stores = await TransferService.GetStoresAsync();

                Filter.DateRangeIndex = 0;
                Filter.DtFrom = DateTime.Today;
                Filter.DtTo = DateTime.Today;

                await LoadDataAsync();
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
                return await Task.FromResult(Makers);
            }

            return await Task.FromResult(Makers.Where(m =>
                m.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
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

        public async Task OnMakerChanged(LookupItemInt? maker)
        {
            SelectedMaker = maker;
            Filter.MakerId = maker?.Id ?? 0;
            await LoadDataAsync();
        }

        public async Task OnArticleChanged(LookupItemString? article)
        {
            SelectedArticle = article;
            Filter.ArticleId = article?.Id ?? string.Empty;
            await LoadDataAsync();
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
                    Filter.DtFrom = DateTime.Today.AddDays(-15);
                    Filter.DtTo = DateTime.Today;
                    break;
                case 2:
                    Filter.DtFrom = DateTime.Today.AddDays(-30);
                    Filter.DtTo = DateTime.Today;
                    break;
                case 3:
                    Filter.DtFrom = DateTime.Today.AddDays(-60);
                    Filter.DtTo = DateTime.Today;
                    break;
                case 4:
                    Filter.DtFrom = DateTime.Today.AddDays(-90);
                    Filter.DtTo = DateTime.Today;
                    break;
                case 5:
                    // Keep user selected custom dates
                    break;
            }

            await LoadDataAsync();
        }

        public async Task OnLotNoKeyDown(KeyboardEventArgs e)
        {
            if (e.Key == "Enter")
            {
                await LoadDataAsync();
            }
        }

        public async Task LoadDataAsync()
        {
            IsLoading = true;
            try
            {
                AvailableLots = await TransferService.GetAvailableLotsAsync(Filter);
                IsAllSelected = false;
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
                IsLoading = false;
            }
        }

        public void ToggleSelectAll(ChangeEventArgs e)
        {
            IsAllSelected = (bool)(e.Value ?? false);
            foreach (var item in AvailableLots)
            {
                item.IsSelected = IsAllSelected;
                if (item.TransferQty <= 0)
                {
                    item.TransferQty = item.AvailQty;
                }
            }
        }

        public void OnSelectionChanged()
        {
            IsAllSelected = AvailableLots.Any() && AvailableLots.All(x => x.IsSelected);
        }

        public void OpenTransferModal()
        {
            if (!SelectedLots.Any())
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "No Lots Selected",
                    Detail = "Please check at least one lot to transfer.",
                    Duration = 4000
                });
                return;
            }

            SelectedStoreId = 0;
            SelectedShelfId = 0;
            TransferRemarks = string.Empty;
            Shelves.Clear();

            ShowTransferModal = true;
        }

        public void CloseTransferModal()
        {
            ShowTransferModal = false;
        }

        public async Task OnStoreChanged()
        {
            if (SelectedStoreId > 0)
            {
                try
                {
                    Shelves = await TransferService.GetShelvesByStoreAsync(SelectedStoreId);
                    SelectedShelfId = 0;
                }
                catch (Exception ex)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Error,
                        Summary = "Error Loading Shelves",
                        Detail = ex.Message,
                        Duration = 4000
                    });
                }
            }
            else
            {
                Shelves.Clear();
                SelectedShelfId = 0;
            }
        }

        public async Task OnShelfChanged()
        {
            if (SelectedShelfId > 0 && SelectedLots.Any())
            {
                string articleId = SelectedArticle?.Id ?? SelectedLots.First().ItemCode;
                try
                {
                    string prevRemarks = await TransferService.GetShelfRemarksAsync(articleId, SelectedShelfId);
                    if (!string.IsNullOrWhiteSpace(prevRemarks))
                    {
                        TransferRemarks = prevRemarks;
                    }
                }
                catch
                {
                    // Non-critical remark prefill failure
                }
            }
        }

        public async Task PostTransfer()
        {
            if (!SelectedLots.Any())
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "No Lots Selected",
                    Detail = "Please select lots to transfer.",
                    Duration = 4000
                });
                return;
            }

            if (SelectedShelfId == 0)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "No Location Selected",
                    Detail = "Please select a warehouse shelf/rack location.",
                    Duration = 4000
                });
                return;
            }

            IsPosting = true;
            try
            {
                string userName = await GetCurrentUserName();
                string machineName = Environment.MachineName;

                var shelf = Shelves.FirstOrDefault(s => s.EntryID == SelectedShelfId);
                string shelfText = shelf?.DisplayName ?? string.Empty;
                string articleId = SelectedArticle?.Id ?? SelectedLots.First().ItemCode;

                var req = new PostTransferRequest
                {
                    ArticleId = articleId,
                    SelectedStoreId = SelectedStoreId,
                    SelectedShelfId = SelectedShelfId,
                    ShelfLocationText = shelfText,
                    Remarks = TransferRemarks,
                    UserName = userName,
                    MachineName = machineName,
                    Items = SelectedLots.Select(x => new TransferItemDto
                    {
                        VRD_EntryID = x.VRD_EntryID,
                        TransferQty = x.TransferQty,
                        OrderNo = x.OrderNo,
                        LotNo = x.LotNo,
                        BatchNo = x.Batch_No,
                        MillCertNo = x.Mill_Certificate_No
                    }).ToList()
                };

                var result = await TransferService.PostTransferToReadyFinishAsync(req);

                if (result.Success)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Success,
                        Summary = "Transfer Posted",
                        Detail = $"Receipt #{result.RcvNo} created for {SelectedLots.Count} lots.",
                        Duration = 5000
                    });

                    ShowTransferModal = false;
                    await LoadDataAsync();
                }
                else
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Error,
                        Summary = "Transfer Failed",
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
                IsPosting = false;
            }
        }

        // ─────────────────────────────────────────────────────────────
        // Reports
        // ─────────────────────────────────────────────────────────────

        public async Task PrintSlip(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem == null) return;

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "PTCQel.rpt",
                SelectionFormula = $"{{VFOrderItemPTC.LotNo}}='{ContextRowItem.LotNo}'"
            });
        }

        public async Task PrintPTC(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem == null) return;

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "PTCQel.rpt",
                SelectionFormula = $"{{VFOrderItemPTC.LotNo}}='{ContextRowItem.LotNo}'"
            });
        }

        private void ResolveRowItem(ItemClickEventArgs args)
        {
            if (args.Data is ReadyFinishStockAvailableItem item)
            {
                ContextRowItem = item;
            }
            else
            {
                ContextRowItem = null;
            }
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
