using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Constants;
using Impulse.Services;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;

namespace Impulse.Pages.Payroll.Designation
{
    public partial class DesignationsList : ComponentBase
    {
        [Inject] private IDesignationService DesignationService { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = default!;
        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;

        [Parameter, SupplyParameterFromQuery]
        public string? ReturnUrl { get; set; }

        private List<DesignationModel> Designations { get; set; } = new();
        private DesignationModel? SelectedRow { get; set; }

        private string SearchText { get; set; } = string.Empty;
        private bool IsLoading { get; set; } = false;
        private bool IsSaving { get; set; } = false;

        // Modal Dialog State
        private bool ShowModal { get; set; } = false;
        private DesignationModel EditModel { get; set; } = new();

        protected override async Task OnInitializedAsync()
        {
            await LoadDesignationsAsync();
        }

        private async Task LoadDesignationsAsync()
        {
            IsLoading = true;
            StateHasChanged();
            try
            {
                Designations = await DesignationService.GetDesignationsAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Designations",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsLoading = false;
                StateHasChanged();
            }
        }

        private IEnumerable<DesignationModel> FilteredList
        {
            get
            {
                if (Designations == null) return Enumerable.Empty<DesignationModel>();
                if (string.IsNullOrWhiteSpace(SearchText)) return Designations;

                return Designations.Where(d =>
                    d.Designation.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                    d.Duties.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                    d.RequiredSkills.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                    d.RequiredTrainings.Contains(SearchText, StringComparison.OrdinalIgnoreCase));
            }
        }

        private void OpenAddModal()
        {
            EditModel = new DesignationModel();
            ShowModal = true;
        }

        private void OpenEditModal(DesignationModel item)
        {
            EditModel = new DesignationModel
            {
                EntryID = item.EntryID,
                Designation = item.Designation,
                Duties = item.Duties,
                RequiredSkills = item.RequiredSkills,
                RequiredTrainings = item.RequiredTrainings
            };
            ShowModal = true;
        }

        private void CloseModal()
        {
            ShowModal = false;
            EditModel = new DesignationModel();
        }

        private async Task SaveDesignationAsync()
        {
            if (string.IsNullOrWhiteSpace(EditModel.Designation))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Error",
                    Detail = "Please enter a Designation Title.",
                    Duration = 3000
                });
                return;
            }

            IsSaving = true;
            StateHasChanged();
            try
            {
                bool success = await DesignationService.SaveDesignationAsync(EditModel);
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Saved",
                        Detail = $"Designation '{EditModel.Designation}' saved successfully.",
                        Duration = 4000
                    });
                    CloseModal();
                    await LoadDesignationsAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Save Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsSaving = false;
                StateHasChanged();
            }
        }

        private async Task DeleteDesignationAsync(DesignationModel item)
        {
            try
            {
                bool success = await DesignationService.DeleteDesignationAsync(item.EntryID);
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Deleted",
                        Detail = $"Designation '{item.Designation}' deleted.",
                        Duration = 3000
                    });
                    await LoadDesignationsAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Delete Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        private async Task ShowOptionsMenu(MouseEventArgs e)
        {
            await BlazorContextMenuService.ShowMenu("optionsBarMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        private async Task ShowRowContextMenu(MouseEventArgs e, DesignationModel item)
        {
            SelectedRow = item;
            await BlazorContextMenuService.ShowMenu("rowContextMenu", (int)e.ClientX, (int)e.ClientY + 15, item);
        }

        private async Task PrintDesignationsList()
        {
            try
            {
                var request = new ReportRequest
                {
                    ReportName = ReportNames.Payroll.DesignationsList
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Report Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        private async Task PrintDesignationDetail(DesignationModel item)
        {
            try
            {
                var request = new ReportRequest
                {
                    ReportName = ReportNames.Payroll.DesignationDetail,
                    SelectionFormula = $"{{Designations.EntryID}}={item.EntryID}"
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Report Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        private void GoBack()
        {
            if (!string.IsNullOrEmpty(ReturnUrl))
            {
                NavigationManager.NavigateTo(ReturnUrl);
            }
            else
            {
                NavigationManager.NavigateTo("/payroll/dashboard");
            }
        }
    }
}
