using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Common;
using DataAccessLibrary.Interface.QMS;
using DataAccessLibrary.Models.QMS;
using Impulse.Services;
using Impulse.Services.Setup;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;

namespace Impulse.Pages.QMS.Calibration
{
    public partial class CalibrationList : ComponentBase
    {
        [Inject] private ICalibrationDataAccess CalibrationData { get; set; } = null!;
        [Inject] private IEntityAttachmentDataAccess AttachmentData { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private NotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IUserPermissionService PermissionService { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNav { get; set; } = null!;

        private List<CalibrationListItemModel> calibrations = new();
        private List<FixedAssetLookupDto> assets = new();
        private CalibrationFilterDto filter = new()
        {
            AssetID = "0",
            DateFrom = DateTime.Today.AddDays(-30),
            DateTo = DateTime.Today
        };

        private string selectedRange = "30";
        private bool isLoading = true;
        private bool showAttachmentsModal = false;
        private CalibrationListItemModel? selectedCalibrationForAttachments;

        protected override async Task OnInitializedAsync()
        {
            isLoading = true;
            try
            {
                await PermissionService.InitializeAsync();
                assets = await CalibrationData.GetFixedAssetsAsync(false);
                await LoadCalibrationsAsync();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Initializing",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                isLoading = false;
            }
        }

        private async Task LoadCalibrationsAsync()
        {
            isLoading = true;
            try
            {
                calibrations = await CalibrationData.GetCalibrationsAsync(filter);
            }
            catch (Exception ex)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Calibrations",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                isLoading = false;
            }
        }

        private async Task OnRangeChanged(ChangeEventArgs e)
        {
            selectedRange = e.Value?.ToString() ?? "30";
            var today = DateTime.Today;

            switch (selectedRange)
            {
                case "0": // Today
                    filter.DateFrom = today;
                    filter.DateTo = today;
                    break;
                case "15":
                    filter.DateFrom = today.AddDays(-15);
                    filter.DateTo = today;
                    break;
                case "30":
                    filter.DateFrom = today.AddDays(-30);
                    filter.DateTo = today;
                    break;
                case "60":
                    filter.DateFrom = today.AddDays(-60);
                    filter.DateTo = today;
                    break;
                case "90":
                    filter.DateFrom = today.AddDays(-90);
                    filter.DateTo = today;
                    break;
                case "180":
                    filter.DateFrom = today.AddDays(-180);
                    filter.DateTo = today;
                    break;
                case "custom":
                    // Keep current dates
                    return;
            }

            await LoadCalibrationsAsync();
        }

        private async Task OnAssetChanged(ChangeEventArgs e)
        {
            filter.AssetID = e.Value?.ToString() ?? "0";
            await LoadCalibrationsAsync();
        }

        private void OnDateFromChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out var dt))
                filter.DateFrom = dt;
            else
                filter.DateFrom = null;
        }

        private void OnDateToChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out var dt))
                filter.DateTo = dt;
            else
                filter.DateTo = null;
        }

        private async Task HandleSearchKeyUp(KeyboardEventArgs e)
        {
            if (e.Key == "Enter")
            {
                await SearchCalibrations();
            }
        }

        private async Task SearchCalibrations()
        {
            await LoadCalibrationsAsync();
        }

        private void CreateNewCalibration()
        {
            NavigationManager.NavigateTo("/qms/calibrations/new?returnUrl=/qms/calibrations");
        }

        private void EditCalibration(CalibrationListItemModel item)
        {
            NavigationManager.NavigateTo($"/qms/calibrations/{item.EntryID}?returnUrl=/qms/calibrations");
        }

        private void OpenAttachments(CalibrationListItemModel item)
        {
            selectedCalibrationForAttachments = item;
            showAttachmentsModal = true;
        }

        private void CloseAttachmentsModal()
        {
            showAttachmentsModal = false;
            selectedCalibrationForAttachments = null;
        }

        private async Task HandleAttachmentCountChanged(int newCount)
        {
            if (selectedCalibrationForAttachments != null)
            {
                selectedCalibrationForAttachments.AttachmentCount = newCount;
            }
            await Task.CompletedTask;
        }

        private async Task DeleteCalibration(CalibrationListItemModel item)
        {
            try
            {
                var success = await CalibrationData.DeleteCalibrationAsync(item.EntryID);
                if (success)
                {
                    calibrations.Remove(item);
                    NotificationServiceManager.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Deleted",
                        Detail = $"Calibration entry #{item.EntryID} deleted.",
                        Duration = 3500
                    });
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Delete Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        private async Task PrintList()
        {
            try
            {
                var fromStr = filter.DateFrom?.ToString("yyyy-MM-dd") ?? DateTime.Today.ToString("yyyy-MM-dd");
                var toStr = filter.DateTo?.ToString("yyyy-MM-dd") ?? DateTime.Today.ToString("yyyy-MM-dd");

                var request = new ReportRequest
                {
                    ReportName = "FixedAssets_CalibrationList.rpt",
                    SelectionFormula = $"{{FixedAssets_Calibration.DT}} in Date({filter.DateFrom?.Year ?? DateTime.Today.Year}, {filter.DateFrom?.Month ?? DateTime.Today.Month}, {filter.DateFrom?.Day ?? DateTime.Today.Day}) to Date({filter.DateTo?.Year ?? DateTime.Today.Year}, {filter.DateTo?.Month ?? DateTime.Today.Month}, {filter.DateTo?.Day ?? DateTime.Today.Day})"
                };

                await ReportNav.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Print Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }
    }
}

