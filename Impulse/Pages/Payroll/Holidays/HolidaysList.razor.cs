using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Constants;
using Impulse.Services;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;

namespace Impulse.Pages.Payroll.Holidays
{
    public partial class HolidaysList : ComponentBase
    {
        [Inject] private IHolidayService HolidayService { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;
        [Inject] private BlazorContextMenu.IBlazorContextMenuService BlazorContextMenuService { get; set; } = default!;

        [Parameter, SupplyParameterFromQuery]
        public string? ReturnUrl { get; set; }

        private List<int> Years { get; set; } = new();
        private int SelectedYear { get; set; } = DateTime.Today.Year;
        private List<HolidayModel> Holidays { get; set; } = new();

        private DateTime NewDate { get; set; } = DateTime.Today;
        private string NewDescription { get; set; } = string.Empty;

        private bool IsLoading { get; set; } = false;
        private bool IsAdding { get; set; } = false;
        private bool IsDeleting { get; set; } = false;

        // Context menu options state
        private bool IsOptionsOpen { get; set; } = false;

        // Edit Modal State
        private bool ShowEditModal { get; set; } = false;
        private HolidayModel? EditingHoliday { get; set; }
        private string EditDescription { get; set; } = string.Empty;

        protected override async Task OnInitializedAsync()
        {
            int currentYear = DateTime.Today.Year;
            for (int y = 2002; y <= currentYear + 2; y++)
            {
                Years.Add(y);
            }
            SelectedYear = currentYear;
            NewDate = new DateTime(SelectedYear, DateTime.Today.Month, DateTime.Today.Day <= DateTime.DaysInMonth(SelectedYear, DateTime.Today.Month) ? DateTime.Today.Day : 1);

            await LoadHolidaysAsync();
        }

        private async Task OnYearChanged(ChangeEventArgs e)
        {
            if (int.TryParse(Convert.ToString(e.Value), out int year))
            {
                SelectedYear = year;
                NewDate = new DateTime(SelectedYear, 1, 1);
                await LoadHolidaysAsync();
            }
        }

        private async Task LoadHolidaysAsync()
        {
            IsLoading = true;
            StateHasChanged();
            try
            {
                Holidays = await HolidayService.GetHolidaysByYearAsync(SelectedYear);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Holidays",
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

        private async Task AddHolidayAsync()
        {
            if (string.IsNullOrWhiteSpace(NewDescription))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Please enter a description for the holiday.",
                    Duration = 3000
                });
                return;
            }

            // Constrain date to selected year
            if (NewDate.Year != SelectedYear)
            {
                NewDate = new DateTime(SelectedYear, NewDate.Month, NewDate.Day);
            }

            IsAdding = true;
            StateHasChanged();
            try
            {
                var model = new HolidayModel
                {
                    DT = NewDate,
                    Description = NewDescription.Trim()
                };

                bool success = await HolidayService.AddHolidayAsync(model);
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Holiday Added",
                        Detail = $"Added holiday '{model.Description}' on {model.DateFormatted}.",
                        Duration = 3000
                    });

                    NewDescription = string.Empty;
                    await LoadHolidaysAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Add Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsAdding = false;
                StateHasChanged();
            }
        }

        private void OpenEditModal(HolidayModel item)
        {
            EditingHoliday = item;
            EditDescription = item.Description;
            ShowEditModal = true;
        }

        private void CloseEditModal()
        {
            ShowEditModal = false;
            EditingHoliday = null;
            EditDescription = string.Empty;
        }

        private async Task SaveEditAsync()
        {
            if (EditingHoliday == null) return;
            if (string.IsNullOrWhiteSpace(EditDescription))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Holiday description cannot be empty.",
                    Duration = 3000
                });
                return;
            }

            StateHasChanged();
            try
            {
                bool success = await HolidayService.UpdateHolidayAsync(EditingHoliday.DT, EditDescription.Trim());
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Holiday Updated",
                        Detail = $"Updated holiday on {EditingHoliday.DateFormatted}.",
                        Duration = 3000
                    });

                    CloseEditModal();
                    await LoadHolidaysAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Update Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        private async Task DeleteHolidayAsync(HolidayModel item)
        {
            IsDeleting = true;
            StateHasChanged();
            try
            {
                bool success = await HolidayService.DeleteHolidayAsync(item.DT);
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Holiday Deleted",
                        Detail = $"Deleted holiday on {item.DateFormatted}.",
                        Duration = 3000
                    });

                    await LoadHolidaysAsync();
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
            finally
            {
                IsDeleting = false;
                StateHasChanged();
            }
        }

        private async Task ShowOptionsMenu(MouseEventArgs e)
        {
            await BlazorContextMenuService.ShowMenu("optionsBarMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        private void ToggleOptions()
        {
            IsOptionsOpen = !IsOptionsOpen;
        }

        private void CloseOptions()
        {
            IsOptionsOpen = false;
        }

        private async Task PrintHolidaysReport()
        {
            CloseOptions();
            try
            {
                var request = new ReportRequest
                {
                    ReportName = ReportNames.Payroll.Holidays,
                    SelectionFormula = $"Year({{Holidays.DT}})={SelectedYear}"
                };
                await ReportNavigationService.PrintReportAsync(request);
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
