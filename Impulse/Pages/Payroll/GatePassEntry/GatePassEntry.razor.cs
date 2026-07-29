using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Payroll.GatePassEntry
{
    public partial class GatePassEntry : ComponentBase
    {
        [Parameter]
        [SupplyParameterFromQuery(Name = "returnUrl")]
        public string? ReturnUrl { get; set; }

        // ── Injected Services ──────────────────────────────────────────────
        [Inject] private IGatePassService     GatePassService     { get; set; } = default!;
        [Inject] private IEmployeeService     EmployeeService     { get; set; } = default!;
        [Inject] private INotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager    NavigationManager   { get; set; } = default!;

        // ── State ──────────────────────────────────────────────────────────
        private bool IsLoading = false;
        private bool IsSaving  = false;

        private DepartmentListItemModel? SelectedDepartment { get; set; }
        private DateTime SelectedDate { get; set; } = DateTime.Today;

        private List<DepartmentListItemModel> AllDepartments = new();
        private List<GatePassRowModel> Rows = new();

        private bool ShowEditModal = false;
        private GatePassRowModel? EditingRow;

        private bool IsAllSelected => Rows.Any() && Rows.All(r => r.IsSelected);

        // ── Initialization ─────────────────────────────────────────────────
        protected override async Task OnInitializedAsync()
        {
            try
            {
                AllDepartments = await EmployeeService.GetDepartmentsAsync(false);
                await LoadListAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Initialization Error", ex.Message);
            }
        }

        // ── Filters & Search ───────────────────────────────────────────────
        private async Task<IEnumerable<DepartmentListItemModel>> SearchDepartments(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return AllDepartments;
            return await Task.FromResult(
                AllDepartments.Where(d => d.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase))
            );
        }

        private async Task OnDepartmentChanged(DepartmentListItemModel? dept)
        {
            SelectedDepartment = dept;
            await LoadListAsync();
        }

        private async Task OnDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e?.Value?.ToString(), out DateTime dt))
            {
                SelectedDate = dt;
                await LoadListAsync();
            }
        }

        private async Task LoadListAsync()
        {
            IsLoading = true;
            StateHasChanged();

            try
            {
                string deptId = SelectedDepartment?.DeptID ?? "0";
                Rows = await GatePassService.GetGatePassListAsync(deptId, SelectedDate);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error Loading Gate Passes", ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
        }

        // ── Selection Controls ─────────────────────────────────────────────
        private void SelectAll()
        {
            foreach (var r in Rows) r.IsSelected = true;
        }

        private void DeselectAll()
        {
            foreach (var r in Rows) r.IsSelected = false;
        }

        private void ToggleSelectAll(ChangeEventArgs e)
        {
            bool select = (bool)(e?.Value ?? false);
            foreach (var r in Rows) r.IsSelected = select;
        }

        // ── Gate Pass Modal Editing & Duration Calculations ───────────────
        private void EditGatePass(GatePassRowModel row)
        {
            EditingRow = new GatePassRowModel
            {
                EntryID    = row.EntryID,
                EmpID      = row.EmpID,
                EmpName    = row.EmpName,
                IsSelected = true,
                OutTime    = row.OutTime,
                InTime     = row.InTime,
                GpHrs      = row.GpHrs,
                IsOfficial = row.IsOfficial,
                Remarks    = row.Remarks
            };
            ShowEditModal = true;
        }

        private void OnModalOutTimeChanged(string? val)
        {
            if (EditingRow == null) return;
            EditingRow.OutTime = FormatTime5(val);
            CalculateDuration(EditingRow);
        }

        private void OnModalInTimeChanged(string? val)
        {
            if (EditingRow == null) return;
            EditingRow.InTime = FormatTime5(val);
            CalculateDuration(EditingRow);
        }

        private static string FormatTime5(string? val)
        {
            if (string.IsNullOrWhiteSpace(val)) return string.Empty;
            val = val.Trim();
            if (TimeSpan.TryParse(val, out TimeSpan ts))
                return ts.ToString(@"hh\:mm");
            if (DateTime.TryParse(val, out DateTime dt))
                return dt.ToString("HH:mm");
            if (val.Length > 5) return val.Substring(0, 5);
            return val;
        }

        private void CalculateDuration(GatePassRowModel row)
        {
            if (TimeSpan.TryParse(row.OutTime, out TimeSpan outTs) && TimeSpan.TryParse(row.InTime, out TimeSpan inTs))
            {
                double dMins = (inTs - outTs).TotalMinutes;
                if (dMins < 0) dMins += 1440; // overnight boundary fallback

                // Friday relaxation logic: If Friday out-time is between 1:00 PM and 2:30 PM
                if (SelectedDate.DayOfWeek == DayOfWeek.Friday && outTs >= new TimeSpan(13, 0, 0) && outTs < new TimeSpan(14, 30, 0))
                {
                    TimeSpan fridayEnd = new TimeSpan(14, 30, 0);
                    if (inTs > fridayEnd)
                    {
                        dMins = (inTs - fridayEnd).TotalMinutes + 30;
                    }
                }

                row.GpHrs = Math.Round(dMins / 60.0, 2);
            }
            else
            {
                row.GpHrs = 0;
            }
        }

        private void ConfirmModalEdit()
        {
            if (EditingRow == null) return;

            var targetRow = Rows.FirstOrDefault(r => r.EmpID == EditingRow.EmpID);
            if (targetRow != null)
            {
                targetRow.IsSelected = true;
                targetRow.OutTime    = FormatTime5(EditingRow.OutTime);
                targetRow.InTime     = FormatTime5(EditingRow.InTime);
                targetRow.GpHrs      = EditingRow.GpHrs;
                targetRow.IsOfficial = EditingRow.IsOfficial;
                targetRow.Remarks    = EditingRow.Remarks;
            }

            ShowEditModal = false;
            EditingRow = null;
        }

        private void CloseModal()
        {
            ShowEditModal = false;
            EditingRow = null;
        }

        // ── Save Action ────────────────────────────────────────────────────
        private async Task SaveAsync()
        {
            IsSaving = true;
            StateHasChanged();

            try
            {
                var input = new GatePassSaveDto
                {
                    Date = SelectedDate,
                    Rows = Rows
                };

                bool success = await GatePassService.SaveGatePassesAsync(input);
                if (success)
                {
                    NotificationService.ShowSuccess("Saved", "Gate pass records saved successfully.");
                    await LoadListAsync();
                }
                else
                {
                    NotificationService.ShowError("Save Failed", "Could not save gate pass records.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Save Error", ex.Message);
            }
            finally
            {
                IsSaving = false;
            }
        }

        private void NavigateBack()
        {
            if (!string.IsNullOrWhiteSpace(ReturnUrl))
                NavigationManager.NavigateTo(ReturnUrl);
            else
                NavigationManager.NavigateTo("/payroll");
        }
    }
}
