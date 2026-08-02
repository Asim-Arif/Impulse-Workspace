using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Components.Authorization;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.Production;
using Impulse.Services;

namespace Impulse.Pages.Production.AuthorizeReceived
{
    public partial class AuthorizeReceived
    {
        [Inject]
        public IAuthorizeReceivedService AuthorizeReceivedService { get; set; } = default!;

        [Inject]
        public NavigationManager NavigationManager { get; set; } = default!;

        [Inject]
        public Radzen.NotificationService NotificationService { get; set; } = default!;

        [Inject]
        public AuthenticationStateProvider AuthenticationStateProvider { get; set; } = default!;

        [Inject]
        public IHttpContextAccessor HttpContextAccessor { get; set; } = default!;

        public AuthorizeReceivedFilterModel Filter { get; set; } = new AuthorizeReceivedFilterModel();
        public List<PendingAuthorizeReceivedItemModel> PendingItems { get; set; } = new List<PendingAuthorizeReceivedItemModel>();

        public List<MakerPOLookupModel> AvailableMakers { get; set; } = new List<MakerPOLookupModel>();
        public List<ProcessPOLookupModel> AvailableProcesses { get; set; } = new List<ProcessPOLookupModel>();
        public List<WastageTypeLookupModel> AvailableWastageTypes { get; set; } = new List<WastageTypeLookupModel>();
        public List<RepairTypeLookupModel> AvailableRepairTypes { get; set; } = new List<RepairTypeLookupModel>();
        public List<EmployeeLookupModel> AvailableInspectors { get; set; } = new List<EmployeeLookupModel>();

        public EmployeeLookupModel? SelectedInspector { get; set; }

        public bool IsLoading { get; set; } = false;
        public bool IsSaving { get; set; } = false;

        public bool AllSelected => PendingItems.Any() && PendingItems.All(i => i.IsChecked);
        public int SelectedCount => PendingItems.Count(i => i.IsChecked);

        // Wastage Modal State
        public bool ShowWastageModal { get; set; } = false;
        public PendingAuthorizeReceivedItemModel? CurrentModalItem { get; set; }
        public List<WastageBreakdownItemModel> TempWastageList { get; set; } = new List<WastageBreakdownItemModel>();

        // Rework Modal State
        public bool ShowReWorkModal { get; set; } = false;
        public List<ReWorkBreakdownItemModel> TempReWorkList { get; set; } = new List<ReWorkBreakdownItemModel>();

        protected override async Task OnInitializedAsync()
        {
            AvailableMakers = await AuthorizeReceivedService.GetMakersAsync();
            AvailableProcesses = await AuthorizeReceivedService.GetProcessesAsync();
            AvailableWastageTypes = await AuthorizeReceivedService.GetWastageTypesAsync();
            AvailableRepairTypes = await AuthorizeReceivedService.GetRepairTypesAsync();
            AvailableInspectors = await AuthorizeReceivedService.GetInspectorsAsync();

            await LoadPendingItemsAsync();
        }

        public async Task LoadPendingItemsAsync()
        {
            IsLoading = true;
            try
            {
                PendingItems = await AuthorizeReceivedService.GetPendingAuthorizeReceivedItemsAsync(Filter);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Load Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        public Task<IEnumerable<EmployeeLookupModel>> SearchInspectors(string searchText)
        {
            if (AvailableInspectors == null || !AvailableInspectors.Any())
                return Task.FromResult(Enumerable.Empty<EmployeeLookupModel>());

            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<EmployeeLookupModel>>(AvailableInspectors);

            return Task.FromResult<IEnumerable<EmployeeLookupModel>>(
                AvailableInspectors.Where(e => (e.Name != null && e.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                                            || (e.EmpID != null && e.EmpID.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                                            || (e.Designation != null && e.Designation.Contains(searchText, StringComparison.OrdinalIgnoreCase))));
        }

        public void SelectAll()
        {
            foreach (var item in PendingItems) item.IsChecked = true;
        }

        public void UnselectAll()
        {
            foreach (var item in PendingItems) item.IsChecked = false;
        }

        public void ToggleSelectAll(ChangeEventArgs e)
        {
            bool check = (bool)(e.Value ?? false);
            foreach (var item in PendingItems) item.IsChecked = check;
        }

        // WASTAGE MODAL HANDLERS
        public void OpenWastageModal(PendingAuthorizeReceivedItemModel item)
        {
            CurrentModalItem = item;
            TempWastageList = item.WastageBreakdowns.Select(w => new WastageBreakdownItemModel
            {
                EntryID = w.EntryID,
                VRD_RefID = w.VRD_RefID,
                Wastage_RefID = w.Wastage_RefID,
                WastageTypeName = w.WastageTypeName,
                Qty = w.Qty,
                WastageType = w.WastageType,
                ReturnTo_VRD_RefID = w.ReturnTo_VRD_RefID,
                VendID = w.VendID,
                VendName = w.VendName,
                EmpID = w.EmpID,
                EmpName = w.EmpName
            }).ToList();

            if (!TempWastageList.Any() && item.Wastage > 0)
            {
                TempWastageList.Add(new WastageBreakdownItemModel
                {
                    VRD_RefID = item.EntryID,
                    Wastage_RefID = AvailableWastageTypes.FirstOrDefault()?.EntryID ?? 0,
                    Qty = item.Wastage,
                    WastageType = 1,
                    VendID = item.VendID
                });
            }

            ShowWastageModal = true;
        }

        public void AddWastageRow()
        {
            if (CurrentModalItem == null) return;
            TempWastageList.Add(new WastageBreakdownItemModel
            {
                VRD_RefID = CurrentModalItem.EntryID,
                Wastage_RefID = AvailableWastageTypes.FirstOrDefault()?.EntryID ?? 0,
                Qty = 1,
                WastageType = 1,
                VendID = CurrentModalItem.VendID
            });
        }

        public void RemoveWastageRow(WastageBreakdownItemModel item)
        {
            TempWastageList.Remove(item);
        }

        public void SaveWastageModal()
        {
            if (CurrentModalItem != null)
            {
                CurrentModalItem.WastageBreakdowns = TempWastageList.ToList();
                CurrentModalItem.Wastage = TempWastageList.Sum(w => w.Qty);
            }
            ShowWastageModal = false;
        }

        public void CloseWastageModal()
        {
            ShowWastageModal = false;
            CurrentModalItem = null;
        }

        // REWORK MODAL HANDLERS
        public void OpenReWorkModal(PendingAuthorizeReceivedItemModel item)
        {
            CurrentModalItem = item;
            TempReWorkList = item.ReWorkBreakdowns.Select(r => new ReWorkBreakdownItemModel
            {
                EntryID = r.EntryID,
                VRD_RefID = r.VRD_RefID,
                Repair_RefID = r.Repair_RefID,
                RepairTypeName = r.RepairTypeName,
                Qty = r.Qty,
                ReturnProcessID = r.ReturnProcessID,
                ReturnProcessName = r.ReturnProcessName
            }).ToList();

            if (!TempReWorkList.Any() && item.ReWorkQty > 0)
            {
                TempReWorkList.Add(new ReWorkBreakdownItemModel
                {
                    VRD_RefID = item.EntryID,
                    Repair_RefID = AvailableRepairTypes.FirstOrDefault()?.EntryID ?? 0,
                    Qty = item.ReWorkQty,
                    ReturnProcessID = item.ProcessID
                });
            }

            ShowReWorkModal = true;
        }

        public void AddReWorkRow()
        {
            if (CurrentModalItem == null) return;
            TempReWorkList.Add(new ReWorkBreakdownItemModel
            {
                VRD_RefID = CurrentModalItem.EntryID,
                Repair_RefID = AvailableRepairTypes.FirstOrDefault()?.EntryID ?? 0,
                Qty = 1,
                ReturnProcessID = CurrentModalItem.ProcessID
            });
        }

        public void RemoveReWorkRow(ReWorkBreakdownItemModel item)
        {
            TempReWorkList.Remove(item);
        }

        public void SaveReWorkModal()
        {
            if (CurrentModalItem != null)
            {
                CurrentModalItem.ReWorkBreakdowns = TempReWorkList.ToList();
                CurrentModalItem.ReWorkQty = TempReWorkList.Sum(r => r.Qty);
            }
            ShowReWorkModal = false;
        }

        public void CloseReWorkModal()
        {
            ShowReWorkModal = false;
            CurrentModalItem = null;
        }

        // POST AUTHORIZATION
        public async Task SaveAuthorization()
        {
            var selectedItems = PendingItems.Where(i => i.IsChecked).ToList();
            if (!selectedItems.Any())
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Selection Required",
                    Detail = "Please select at least 1 entry to authorize.",
                    Duration = 4000
                });
                return;
            }

            IsSaving = true;
            try
            {
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                string userName = authState?.User?.Identity?.Name ?? "System";
                int userId = 1;

                var httpContext = HttpContextAccessor.HttpContext;
                string machineName = "127.0.0.1";
                if (httpContext != null)
                {
                    string? ip = httpContext.Request.Headers["X-Forwarded-For"].FirstOrDefault();
                    if (string.IsNullOrEmpty(ip))
                    {
                        ip = httpContext.Connection.RemoteIpAddress?.ToString();
                    }
                    if (!string.IsNullOrEmpty(ip))
                    {
                        machineName = ip;
                    }
                    else
                    {
                        machineName = Environment.MachineName;
                    }
                }

                string inspectorEmpId = SelectedInspector?.EmpID ?? "";

                bool success = await AuthorizeReceivedService.SaveAuthorizationAsync(selectedItems, inspectorEmpId, userName, machineName, userId);

                if (success)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Success,
                        Summary = "Authorization Posted",
                        Detail = $"Successfully authorized {selectedItems.Count} received lot entry(ies).",
                        Duration = 5000
                    });

                    await LoadPendingItemsAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Authorization Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsSaving = false;
            }
        }
    }
}
