using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using BlazorContextMenu;
using DataAccessLibrary.Interface.QMS;
using DataAccessLibrary.Models.QMS;
using Impulse.Components.Common;
using Impulse.Constants;
using Impulse.Models;
using Impulse.Services;
using Impulse.Services.Setup;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;

namespace Impulse.Pages.QMS.CustomerComplaints
{
    public partial class CustomerComplaintsList : ComponentBase
    {
        [Inject] private ICustomerComplaintDataAccess ComplaintData { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private NotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private IUserPermissionService PermissionService { get; set; } = null!;

        private List<CustomerComplaintListItemModel> complaints = new();
        private List<string> customerList = new();
        private List<string> countries = new();

        private string? selectedCustomer = null;
        private string selectedCountry = string.Empty;
        private int selectedRangeIndex = 0;
        private DateTime? dateFrom = DateTime.Today;
        private DateTime? dateTo = DateTime.Today;
        private string complaintNoSearch = string.Empty;

        private bool includeComplaints = true;
        private bool includeFeedback = true;
        private bool includeAudits = true;

        private bool isLoading = true;
        private bool isAllSelected = false;

        private AttachmentModal? attachmentModalRef;
        private bool showCapaReportModal = false;
        private DateTime capaReportFrom = DateTime.Today.AddDays(-30);
        private DateTime capaReportTo = DateTime.Today;

        private int SelectedCount => complaints.Count(c => c.IsSelected);

        protected override async Task OnInitializedAsync()
        {
            try
            {
                await PermissionService.InitializeAsync();
                await LoadLookupsAsync();
                await LoadComplaintsAsync();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Initialization Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                isLoading = false;
            }
        }

        private async Task LoadLookupsAsync()
        {
            customerList = await ComplaintData.GetCustomerCodesAsync();
            countries = await ComplaintData.GetCountriesAsync();
        }

        private async Task<IEnumerable<string>> SearchCustomers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return customerList;
            }

            return await Task.FromResult(
                customerList.Where(c => c.Contains(searchText, StringComparison.OrdinalIgnoreCase))
            );
        }

        private async Task OnCustomerChanged(string? customer)
        {
            selectedCustomer = customer;
            countries = await ComplaintData.GetCountriesAsync(selectedCustomer);
            if (!countries.Contains(selectedCountry))
            {
                selectedCountry = string.Empty;
            }
            await LoadComplaintsAsync();
        }

        private async Task OnCountryChanged(ChangeEventArgs e)
        {
            selectedCountry = e.Value?.ToString() ?? string.Empty;
            await LoadComplaintsAsync();
        }

        private async Task OnDateRangePresetChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int rangeIdx))
            {
                selectedRangeIndex = rangeIdx;
                var today = DateTime.Today;

                switch (selectedRangeIndex)
                {
                    case 0: // Today
                        dateFrom = today;
                        dateTo = today;
                        break;
                    case 1: // Last 15 Days
                        dateFrom = today.AddDays(-15);
                        dateTo = today;
                        break;
                    case 2: // Last 30 Days
                        dateFrom = today.AddDays(-30);
                        dateTo = today;
                        break;
                    case 3: // Last 60 Days
                        dateFrom = today.AddDays(-60);
                        dateTo = today;
                        break;
                    case 4: // Last 90 Days
                        dateFrom = today.AddDays(-90);
                        dateTo = today;
                        break;
                    case 5: // Last 180 Days
                        dateFrom = today.AddDays(-180);
                        dateTo = today;
                        break;
                    case 6: // Custom
                        // Keep current dates
                        break;
                }

                await LoadComplaintsAsync();
            }
        }

        private async Task OnDateFromChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out var dt))
            {
                dateFrom = dt;
                await LoadComplaintsAsync();
            }
        }

        private async Task OnDateToChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out var dt))
            {
                dateTo = dt;
                await LoadComplaintsAsync();
            }
        }

        private async Task HandleSearchKeyUp(KeyboardEventArgs e)
        {
            if (e.Key == "Enter")
            {
                await LoadComplaintsAsync();
            }
        }

        private async Task OnToggleComplaints(ChangeEventArgs e)
        {
            includeComplaints = (bool)(e.Value ?? true);
            await LoadComplaintsAsync();
        }

        private async Task OnToggleFeedback(ChangeEventArgs e)
        {
            includeFeedback = (bool)(e.Value ?? true);
            await LoadComplaintsAsync();
        }

        private async Task OnToggleAudits(ChangeEventArgs e)
        {
            includeAudits = (bool)(e.Value ?? true);
            await LoadComplaintsAsync();
        }

        private async Task LoadComplaintsAsync()
        {
            isLoading = true;
            try
            {
                var filter = new CustomerComplaintFilterDto
                {
                    CustCode = selectedCustomer,
                    Country = selectedCountry,
                    ComplaintNo = complaintNoSearch,
                    DateFrom = dateFrom,
                    DateTo = dateTo,
                    IncludeComplaints = includeComplaints,
                    IncludeFeedback = includeFeedback,
                    IncludeAudits = includeAudits
                };

                complaints = await ComplaintData.GetComplaintsAsync(filter);
                isAllSelected = false;
            }
            catch (Exception ex)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Failed to load entries",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                isLoading = false;
            }
        }

        private void ToggleSelectAll(ChangeEventArgs e)
        {
            isAllSelected = (bool)(e.Value ?? false);
            foreach (var item in complaints)
            {
                item.IsSelected = isAllSelected;
            }
        }

        private void OnItemSelectionChanged(CustomerComplaintListItemModel item, ChangeEventArgs e)
        {
            item.IsSelected = (bool)(e.Value ?? false);
            isAllSelected = complaints.Count > 0 && complaints.All(c => c.IsSelected);
        }

        private async Task OpenOptionsMenu(MouseEventArgs e)
        {
            await BlazorContextMenuService.ShowMenu("optionsBarMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        private async Task OpenManageMenu(MouseEventArgs e, CustomerComplaintListItemModel item)
        {
            await BlazorContextMenuService.ShowMenu("rowContextMenu", (int)e.ClientX, (int)e.ClientY + 15, item);
        }

        private CustomerComplaintListItemModel? highlightedComplaint;

        private void OpenAttachments(CustomerComplaintListItemModel item)
        {
            highlightedComplaint = item;
            attachmentModalRef?.Open(
                entityType: "CustomerComplaint",
                entityId: item.EntryID.ToString(),
                secondaryId: null,
                displayName: $"[{item.ComplaintNo}] {item.CustCode} - {item.Country}"
            );
        }

        private void HandleAttachmentCountChanged(int count)
        {
            if (highlightedComplaint != null)
            {
                highlightedComplaint.AttachmentCount = count;
                StateHasChanged();
            }
        }

        private void NavigateToHub()
        {
            NavigationManager.NavigateTo("/qms");
        }

        private void AddNewComplaint()
        {
            NavigationManager.NavigateTo("/qms/customer-complaint/new?returnUrl=/qms/command-center");
        }

        private void AddNewFeedback()
        {
            NavigationManager.NavigateTo("/qms/customer-feedback/new?returnUrl=/qms/command-center");
        }

        private void AddNewAudit()
        {
            NavigationManager.NavigateTo("/qms/audit/new?returnUrl=/qms/command-center");
        }

        private void EditComplaint(CustomerComplaintListItemModel item)
        {
            var returnUrl = "/qms/command-center";
            var url = item.Type switch
            {
                1 => $"/qms/customer-feedback/{item.EntryID}?returnUrl={returnUrl}",
                2 => $"/qms/audit/{item.EntryID}?returnUrl={returnUrl}",
                _ => $"/qms/customer-complaint/{item.EntryID}?returnUrl={returnUrl}"
            };

            NavigationManager.NavigateTo(url);
        }

        private async Task GenerateNC(CustomerComplaintListItemModel item)
        {
            var existingNc = await ComplaintData.GetNCNoAsync(item.EntryID);
            if (!string.IsNullOrWhiteSpace(existingNc))
            {
                var ncId = await ComplaintData.GetNCEntryIdAsync(existingNc);
                if (ncId > 0)
                {
                    NavigationManager.NavigateTo($"/qms/non-compliance/{ncId}?complaintId={item.EntryID}&returnUrl=/qms/command-center");
                    return;
                }
            }

            NavigationManager.NavigateTo($"/qms/non-compliance/new?complaintId={item.EntryID}&returnUrl=/qms/command-center");
        }

        private async Task OpenNC(CustomerComplaintListItemModel item)
        {
            if (string.IsNullOrWhiteSpace(item.NCNo)) return;
            var ncId = await ComplaintData.GetNCEntryIdAsync(item.NCNo);
            if (ncId > 0)
            {
                NavigationManager.NavigateTo($"/qms/non-compliance/{ncId}?complaintId={item.EntryID}&returnUrl=/qms/command-center");
            }
        }

        private async Task GenerateCAPA(CustomerComplaintListItemModel item)
        {
            var ncNo = await ComplaintData.GetNCNoAsync(item.EntryID);
            if (string.IsNullOrWhiteSpace(ncNo))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "NC Required",
                    Detail = "CAPA can only be generated against an existing NC.",
                    Duration = 4000
                });
                return;
            }

            var capaNo = await ComplaintData.GetCAPANoAsync(item.EntryID);
            if (!string.IsNullOrWhiteSpace(capaNo))
            {
                var capaId = await ComplaintData.GetCAPAEntryIdAsync(capaNo);
                NavigationManager.NavigateTo($"/qms/capa/{capaId}?complaintId={item.EntryID}&returnUrl=/qms/command-center");
            }
            else
            {
                var ncId = await ComplaintData.GetNCEntryIdAsync(ncNo);
                NavigationManager.NavigateTo($"/qms/capa/new?complaintId={item.EntryID}&ncId={ncId}&returnUrl=/qms/command-center");
            }
        }

        private void GenerateCreditNote(CustomerComplaintListItemModel item)
        {
            NavigationManager.NavigateTo($"/qms/credit-note/{item.EntryID}?complaintId={item.EntryID}&returnUrl=/qms/command-center");
        }

        private void OpenCreditNote(CustomerComplaintListItemModel item)
        {
            NavigationManager.NavigateTo($"/qms/credit-note/{item.EntryID}?complaintId={item.EntryID}&returnUrl=/qms/command-center");
        }

        private async Task PrintEntry(CustomerComplaintListItemModel item)
        {
            var request = new ReportRequest
            {
                ReportName = ReportNames.QMS.CustomerComplaint,
                SelectionFormula = $"{{CustomerComplaints.EntryID}}={item.EntryID}"
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintCAPA(CustomerComplaintListItemModel item)
        {
            if (string.IsNullOrWhiteSpace(item.CAPANo))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "No CAPA",
                    Detail = "No CAPA number assigned to this entry.",
                    Duration = 3000
                });
                return;
            }

            var request = new ReportRequest
            {
                ReportName = ReportNames.QMS.CAPA,
                SelectionFormula = $"{{CAPA.CAPANo}}='{item.CAPANo}'"
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintCreditNote(CustomerComplaintListItemModel item)
        {
            if (string.IsNullOrWhiteSpace(item.CreditNoteNo))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "No Credit Note",
                    Detail = "No Credit Note number assigned to this entry.",
                    Duration = 3000
                });
                return;
            }

            var request = new ReportRequest
            {
                ReportName = ReportNames.QMS.CreditNote,
                SelectionFormula = $"{{CustomerComplaints.EntryID}}={item.EntryID} AND {{CustomerComplaints_Detail.Price}}<>0"
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintSelectedList()
        {
            var selectedIds = complaints.Where(c => c.IsSelected).Select(c => c.EntryID).ToList();
            if (selectedIds.Count == 0)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Selection Required",
                    Detail = "Please select at least one entry to print.",
                    Duration = 3000
                });
                return;
            }

            var request = new ReportRequest
            {
                ReportName = ReportNames.QMS.CustomerComplaintsList,
                SelectionFormula = $"{{CustomerComplaints.EntryID}} IN [{string.Join(",", selectedIds)}]"
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintSelectedListWithoutPrice()
        {
            var selectedIds = complaints.Where(c => c.IsSelected).Select(c => c.EntryID).ToList();
            if (selectedIds.Count == 0)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Selection Required",
                    Detail = "Please select at least one entry to print.",
                    Duration = 3000
                });
                return;
            }

            var request = new ReportRequest
            {
                ReportName = ReportNames.QMS.CustomerComplaintsListWithoutPrice,
                SelectionFormula = $"{{CustomerComplaints.EntryID}} IN [{string.Join(",", selectedIds)}]"
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        private void OpenCapaReportModal()
        {
            capaReportFrom = DateTime.Today.AddDays(-30);
            capaReportTo = DateTime.Today;
            showCapaReportModal = true;
        }

        private async Task PrintCapaReportConfirm()
        {
            showCapaReportModal = false;
            var request = new ReportRequest
            {
                ReportName = ReportNames.QMS.CAPAReport,
                SelectionFormula = $"{{CAPA.DT}} IN Date({capaReportFrom.Year},{capaReportFrom.Month},{capaReportFrom.Day}) TO Date({capaReportTo.Year},{capaReportTo.Month},{capaReportTo.Day})"
            };
            await ReportNavigation.PrintReportAsync(request);
        }
    }
}

