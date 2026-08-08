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

namespace Impulse.Pages.Production.CreateDispatchList
{
    public partial class CreateDispatchList : ComponentBase
    {
        [Inject] private ICreateDispatchListService DispatchListService { get; set; } = default!;
        [Inject] private Radzen.NotificationService NotificationService { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = default!;

        // Lookups
        public List<LookupItemString> Customers { get; set; } = new List<LookupItemString>();
        public List<string> Countries { get; set; } = new List<string>();
        public List<LookupItemInt> Makers { get; set; } = new List<LookupItemInt>();
        public List<LookupItemString> Articles { get; set; } = new List<LookupItemString>();

        // Grid Data
        public List<FinalizedLotItem> AvailableLots { get; set; } = new List<FinalizedLotItem>();
        public List<DispatchListStagedItem> StagedItems { get; set; } = new List<DispatchListStagedItem>();

        // Selections & Filters
        public LookupItemString? SelectedCustomer { get; set; } = null;
        public LookupItemInt? SelectedMaker { get; set; } = null;
        public LookupItemString? SelectedArticle { get; set; } = null;
        public DispatchListFilter Filter { get; set; } = new DispatchListFilter();

        // State & Tabs
        [Parameter] [SupplyParameterFromQuery] public long? EntryId { get; set; }

        public int ActiveTab { get; set; } = 0; // 0 = Lots Finalyzed, 1 = Dispatch List
        public string DispatchListNo { get; set; } = string.Empty;
        public long SavedEntryId { get; set; } = 0;
        public List<long> DeletedInnerIDs { get; set; } = new List<long>();

        public bool IsLoading { get; set; } = false;
        public bool IsSaving { get; set; } = false;
        public bool IsAllSelected { get; set; } = false;
        public bool IsAllStagedSelected { get; set; } = false;

        // Box Packaging Modal State (Tab 0 -> Tab 1)
        public bool ShowStageModal { get; set; } = false;
        public decimal StageQtyPerBox { get; set; } = 50;

        // Add to Carton Modal State (Tab 1)
        public bool ShowCartonModal { get; set; } = false;
        public int AssignedCartonNo { get; set; } = 1;
        public DispatchListStagedItem? ContextRowItem { get; set; } = null;

        // Computed Properties
        public List<FinalizedLotItem> SelectedLots => AvailableLots.Where(x => x.IsSelected).ToList();
        public List<DispatchListStagedItem> SelectedStagedItems => StagedItems.Where(x => x.IsSelected).ToList();

        public decimal TotalAvailableQty => AvailableLots.Sum(x => x.AvailQty);
        public decimal TotalSelectedQty => SelectedLots.Sum(x => x.DispatchQty);
        public decimal TotalStagedQty => StagedItems.Sum(x => x.Qty);
        public int TotalCartonsCount => StagedItems.Where(x => x.CartonNo > 0).Select(x => x.CartonNo).Distinct().Count();

        protected override async Task OnInitializedAsync()
        {
            try
            {
                Customers = await DispatchListService.GetCustomersAsync();
                Makers = await DispatchListService.GetMakersAsync();
                Articles = await DispatchListService.GetArticlesAsync();

                Filter.DateRangeIndex = 0;
                Filter.DtFrom = DateTime.Today;
                Filter.DtTo = DateTime.Today;

                if (EntryId.HasValue && EntryId.Value > 0)
                {
                    var loaded = await DispatchListService.GetDispatchListForEditAsync(EntryId.Value);
                    if (loaded != null)
                    {
                        SavedEntryId = loaded.EntryID;
                        DispatchListNo = loaded.DispatchListNo;
                        Filter.CustCode = loaded.CustCode;
                        Filter.Country = loaded.Country;
                        SelectedCustomer = Customers.FirstOrDefault(c => c.Id == loaded.CustCode);
                        StagedItems = loaded.Items;
                        ActiveTab = 1; // Open directly to Dispatch List tab

                        if (!string.IsNullOrEmpty(Filter.CustCode))
                        {
                            Countries = await DispatchListService.GetCountriesByCustomerAsync(Filter.CustCode);
                        }

                        NotificationService.Notify(new Radzen.NotificationMessage
                        {
                            Severity = Radzen.NotificationSeverity.Info,
                            Summary = "Editing Dispatch List",
                            Detail = $"Loaded '{DispatchListNo}' with {StagedItems.Count} box lines.",
                            Duration = 3500
                        });
                    }
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


        public async Task<IEnumerable<LookupItemString>> SearchCustomers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return await Task.FromResult(Customers);
            }

            return await Task.FromResult(Customers.Where(c =>
                c.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
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

        public async Task OnCustomerChanged(LookupItemString? cust)
        {
            SelectedCustomer = cust;
            Filter.CustCode = cust?.Id ?? string.Empty;
            Countries.Clear();
            Filter.Country = string.Empty;

            if (!string.IsNullOrEmpty(Filter.CustCode))
            {
                Countries = await DispatchListService.GetCountriesByCustomerAsync(Filter.CustCode);
                if (Countries.Any())
                {
                    Filter.Country = Countries.First();
                }
            }

            await LoadLotsAsync();
        }

        public async Task OnMakerChanged(LookupItemInt? maker)
        {
            SelectedMaker = maker;
            Filter.MakerId = maker?.Id ?? 0;
            await LoadLotsAsync();
        }

        public async Task OnArticleChanged(LookupItemString? article)
        {
            SelectedArticle = article;
            Filter.ArticleId = article?.Id ?? string.Empty;
            await LoadLotsAsync();
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
            }

            await LoadLotsAsync();
        }

        public async Task OnLotNoKeyDown(KeyboardEventArgs e)
        {
            if (e.Key == "Enter")
            {
                await LoadLotsAsync();
            }
        }

        public async Task LoadLotsAsync()
        {
            if (string.IsNullOrWhiteSpace(Filter.CustCode))
            {
                AvailableLots.Clear();
                return;
            }

            IsLoading = true;
            try
            {
                AvailableLots = await DispatchListService.GetAvailableFinalizedLotsAsync(Filter);
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

        public void SetTab(int tab)
        {
            ActiveTab = tab;
        }

        public void ToggleSelectAll(ChangeEventArgs e)
        {
            IsAllSelected = (bool)(e.Value ?? false);
            foreach (var item in AvailableLots)
            {
                item.IsSelected = IsAllSelected;
                if (item.DispatchQty <= 0)
                {
                    item.DispatchQty = item.AvailQty;
                }
            }
        }

        public void OnSelectionChanged()
        {
            IsAllSelected = AvailableLots.Any() && AvailableLots.All(x => x.IsSelected);
        }

        public void OpenStageModal()
        {
            if (!SelectedLots.Any())
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "No Lots Selected",
                    Detail = "Please select at least one lot to stage.",
                    Duration = 4000
                });
                return;
            }

            if (TotalSelectedQty < 50)
            {
                StageQtyPerBox = TotalSelectedQty;
            }
            else if (StageQtyPerBox <= 0)
            {
                StageQtyPerBox = 50;
            }

            ShowStageModal = true;
        }

        public void CloseStageModal()
        {
            ShowStageModal = false;
        }

        public void ConfirmStagingBoxes()
        {
            if (StageQtyPerBox <= 0)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Invalid Qty / Box",
                    Detail = "Please specify a positive quantity per box.",
                    Duration = 4000
                });
                return;
            }

            var lotQueue = new Queue<FinalizedLotItem>(SelectedLots.Where(x => x.DispatchQty > 0));
            FinalizedLotItem? currentLot = lotQueue.Count > 0 ? lotQueue.Dequeue() : null;
            decimal currentLotRemaining = currentLot?.DispatchQty ?? 0;

            int addedCount = 0;
            while (currentLot != null)
            {
                decimal boxNeeded = StageQtyPerBox;

                while (boxNeeded > 0 && currentLot != null)
                {
                    decimal takeQty = Math.Min(boxNeeded, currentLotRemaining);

                    StagedItems.Add(new DispatchListStagedItem
                    {
                        VRD_EntryID = currentLot.VRD_EntryID,
                        Qty = takeQty,
                        ItemCode = currentLot.ItemCode,
                        ItemName = currentLot.ItemName,
                        ProcessID = currentLot.ProcessID,
                        ProcessDescription = currentLot.Description,
                        CartonNo = 0, // Unassigned until Add to Carton
                        LotNo = currentLot.LotNo,
                        OrderNo = currentLot.OrderNo,
                        IsSelected = false
                    });

                    addedCount++;
                    boxNeeded -= takeQty;
                    currentLotRemaining -= takeQty;

                    if (currentLotRemaining <= 0)
                    {
                        currentLot = lotQueue.Count > 0 ? lotQueue.Dequeue() : null;
                        currentLotRemaining = currentLot?.DispatchQty ?? 0;
                    }
                }
            }

            ShowStageModal = false;
            NotificationService.Notify(new Radzen.NotificationMessage
            {
                Severity = Radzen.NotificationSeverity.Success,
                Summary = "Boxes Staged",
                Detail = $"{addedCount} box lines created on Dispatch List.",
                Duration = 3500
            });

            ActiveTab = 1;
        }

        public void ToggleSelectAllStaged(ChangeEventArgs e)
        {
            IsAllStagedSelected = (bool)(e.Value ?? false);
            foreach (var s in StagedItems)
            {
                s.IsSelected = IsAllStagedSelected;
            }
        }

        public void OnStagedSelectionChanged()
        {
            IsAllStagedSelected = StagedItems.Any() && StagedItems.All(x => x.IsSelected);
        }

        public void OpenCartonModal()
        {
            if (!SelectedStagedItems.Any())
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "No Staged Boxes Selected",
                    Detail = "Please check one or more boxes to assign to a carton.",
                    Duration = 4000
                });
                return;
            }

            ContextRowItem = null;
            int maxCarton = StagedItems.Where(x => x.CartonNo > 0).Select(x => x.CartonNo).DefaultIfEmpty(0).Max();
            AssignedCartonNo = maxCarton + 1;
            ShowCartonModal = true;
        }

        public void CloseCartonModal()
        {
            ShowCartonModal = false;
            ContextRowItem = null;
        }

        public void ApplyCartonNo()
        {
            if (AssignedCartonNo <= 0)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Invalid Carton Number",
                    Detail = "Please enter a carton number greater than 0.",
                    Duration = 4000
                });
                return;
            }

            var targets = ContextRowItem != null
                ? new List<DispatchListStagedItem> { ContextRowItem }
                : SelectedStagedItems;

            if (!targets.Any())
            {
                targets = SelectedStagedItems;
            }

            foreach (var s in targets)
            {
                s.CartonNo = AssignedCartonNo;
            }

            int count = targets.Count;
            ShowCartonModal = false;
            ContextRowItem = null;

            NotificationService.Notify(new Radzen.NotificationMessage
            {
                Severity = Radzen.NotificationSeverity.Success,
                Summary = "Carton Assigned",
                Detail = $"Carton {AssignedCartonNo} assigned to {count} box row(s).",
                Duration = 3000
            });
        }

        public void ContextAddToCarton(ItemClickEventArgs args)
        {
            if (args.Data is DispatchListStagedItem item)
            {
                if (item.IsSelected && SelectedStagedItems.Count > 1)
                {
                    ContextRowItem = null;
                }
                else
                {
                    item.IsSelected = true;
                    ContextRowItem = item;
                }

                int maxCarton = StagedItems.Where(x => x.CartonNo > 0).Select(x => x.CartonNo).DefaultIfEmpty(0).Max();
                AssignedCartonNo = item.CartonNo > 0 ? item.CartonNo : maxCarton + 1;
                ShowCartonModal = true;
            }
        }

        public void ContextRemoveRow(ItemClickEventArgs args)
        {
            if (args.Data is DispatchListStagedItem item)
            {
                if (item.InnerEntryID > 0)
                {
                    DeletedInnerIDs.Add(item.InnerEntryID);
                }
                StagedItems.Remove(item);
                IsAllStagedSelected = false;
            }
        }

        public void RemoveSelectedStaged()
        {
            foreach (var s in StagedItems.Where(x => x.IsSelected))
            {
                if (s.InnerEntryID > 0)
                {
                    DeletedInnerIDs.Add(s.InnerEntryID);
                }
            }
            StagedItems.RemoveAll(x => x.IsSelected);
            IsAllStagedSelected = false;
        }

        public async Task SaveDispatchList()
        {
            if (string.IsNullOrWhiteSpace(DispatchListNo))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Missing Dispatch List #",
                    Detail = "Please enter a Dispatch List Number before saving.",
                    Duration = 4000
                });
                return;
            }

            if (!StagedItems.Any())
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "No Items Staged",
                    Detail = "Please stage items in the dispatch list first.",
                    Duration = 4000
                });
                return;
            }

            IsSaving = true;
            try
            {
                string userName = await GetCurrentUserName();
                string machineName = Environment.MachineName;

                var req = new SaveDispatchListRequest
                {
                    EntryID = SavedEntryId,
                    DispatchListNo = DispatchListNo.Trim(),
                    CustCode = Filter.CustCode,
                    Country = Filter.Country,
                    Items = StagedItems,
                    DeletedInnerIDs = DeletedInnerIDs,
                    UserName = userName,
                    MachineName = machineName
                };

                var result = await DispatchListService.SaveDispatchListAsync(req);

                if (result.Success)
                {
                    SavedEntryId = result.DispatchListEntryID;
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Success,
                        Summary = "Dispatch List Saved",
                        Detail = result.Message,
                        Duration = 5000
                    });

                    await LoadLotsAsync();
                }
                else
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Error,
                        Summary = "Save Failed",
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

        // Manual Qty Modal State
        public bool ShowManualQtyModal { get; set; } = false;
        public decimal ManualQtyInput { get; set; } = 50;
        public bool PendingOtherLabel { get; set; } = false;

        public async Task PrintInnerLabels(bool bOtherLabel = false, bool bWithManualQty = false)
        {
            var targetRows = SelectedStagedItems.Any()
                ? SelectedStagedItems
                : (ContextRowItem != null ? new List<DispatchListStagedItem> { ContextRowItem } : new List<DispatchListStagedItem>());

            if (!targetRows.Any())
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "No Boxes Selected",
                    Detail = "Please select one or more box rows using the checkboxes or right-click to print inner labels.",
                    Duration = 4000
                });
                return;
            }

            if (targetRows.Any(x => x.InnerEntryID == 0))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Unsaved Changes",
                    Detail = "Please save the dispatch list first before printing inner labels.",
                    Duration = 4000
                });
                return;
            }

            if (bWithManualQty)
            {
                PendingOtherLabel = bOtherLabel;
                ManualQtyInput = targetRows.First().Qty;
                ShowManualQtyModal = true;
            }
            else
            {
                await ExecutePrintInnerLabels(bOtherLabel, false, null);
            }
        }

        public void CloseManualQtyModal()
        {
            ShowManualQtyModal = false;
        }

        public async Task ConfirmManualQtyPrint()
        {
            ShowManualQtyModal = false;
            await ExecutePrintInnerLabels(PendingOtherLabel, true, ManualQtyInput);
        }

        private async Task ExecutePrintInnerLabels(bool bOtherLabel, bool bWithManualQty, decimal? manualQty)
        {
            try
            {
                var targetRows = SelectedStagedItems.Any()
                    ? SelectedStagedItems
                    : (ContextRowItem != null ? new List<DispatchListStagedItem> { ContextRowItem } : new List<DispatchListStagedItem>());

                if (!targetRows.Any()) return;

                string? reportName = await DispatchListService.GetInnerLabelReportNameAsync(
                    Filter.CustCode,
                    Filter.Country,
                    otherLabel: bOtherLabel,
                    manualQty: bWithManualQty);

                if (string.IsNullOrWhiteSpace(reportName))
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Warning,
                        Summary = "Report Template Missing",
                        Detail = $"Inner label report template is not defined in Foreign Customers for Customer '{Filter.CustCode}' and Country '{Filter.Country}'.",
                        Duration = 6000
                    });
                    return;
                }

                string strEntries = string.Join(",", targetRows.Select(x => x.InnerEntryID));
                var reportReq = new ReportRequest
                {
                    ReportName = reportName,
                    SelectionFormula = $"{{VPrintInnerLabels.EntryID}} in [{strEntries}]"
                };

                if (bWithManualQty && manualQty.HasValue)
                {
                    reportReq.Parameters["QtyToShow"] = manualQty.Value.ToString();
                }

                await ReportNavigationService.PrintReportAsync(reportReq);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Print Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        public async Task PrintDispatchListReport()
        {
            if (SavedEntryId == 0) return;

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "DispatchList.rpt",
                SelectionFormula = $"{{DispatchList.EntryID}} = {SavedEntryId}"
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
