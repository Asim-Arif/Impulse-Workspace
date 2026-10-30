using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Common;
using DataAccessLibrary.Interface.QMS;
using DataAccessLibrary.Models.QMS;
using Impulse.Services.Setup;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.QMS.Calibration
{
    public partial class CalibrationEntry : ComponentBase
    {
        [Parameter] public int Id { get; set; }
        [SupplyParameterFromQuery(Name = "returnUrl")] public string? ReturnUrl { get; set; }

        [Inject] private ICalibrationDataAccess CalibrationData { get; set; } = null!;
        [Inject] private IEntityAttachmentDataAccess AttachmentData { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private NotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IUserPermissionService PermissionService { get; set; } = null!;

        private bool IsEditMode => Id > 0;
        private string activeTab = "details";
        private bool isLoading = true;
        private bool isSaving = false;
        private int attachmentCount = 0;

        private CalibrationModel model = new();
        private List<FixedAssetLookupDto> assets = new();
        private List<VendorLookupDto> vendors = new();

        private FixedAssetLookupDto? selectedAsset;
        private VendorLookupDto? selectedVendor;

        protected override async Task OnInitializedAsync()
        {
            isLoading = true;
            try
            {
                await PermissionService.InitializeAsync();
                assets = await CalibrationData.GetFixedAssetsAsync(false);
                vendors = await CalibrationData.GetVendorsAsync();

                if (IsEditMode)
                {
                    var existing = await CalibrationData.GetCalibrationByIdAsync(Id);
                    if (existing != null)
                    {
                        model = existing;
                        selectedAsset = assets.FirstOrDefault(a => a.AssetID.Equals(model.AssetID, StringComparison.OrdinalIgnoreCase));
                        if (!string.IsNullOrWhiteSpace(model.AccNo))
                        {
                            selectedVendor = vendors.FirstOrDefault(v => v.AccNo.Equals(model.AccNo, StringComparison.OrdinalIgnoreCase));
                        }

                        var atts = await AttachmentData.GetAttachmentsAsync("Calibration", Id.ToString());
                        attachmentCount = atts.Count;
                    }
                    else
                    {
                        NotificationServiceManager.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Error,
                            Summary = "Not Found",
                            Detail = $"Calibration entry #{Id} could not be found.",
                            Duration = 4000
                        });
                        Cancel();
                        return;
                    }
                }
                else
                {
                    model = new CalibrationModel
                    {
                        DT = DateTime.Today,
                        NextDueDT = DateTime.Today.AddDays(7),
                        CalibrationFrequency = 0
                    };
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Initializing Form",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                isLoading = false;
            }
        }

        private void SetActiveTab(string tab)
        {
            activeTab = tab;
        }

        private Task<IEnumerable<FixedAssetLookupDto>> SearchAssets(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<FixedAssetLookupDto>>(assets.Take(50));

            return Task.FromResult(assets.Where(a =>
                a.AssetID.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                a.Description.Contains(searchText, StringComparison.OrdinalIgnoreCase)
            ).Take(50));
        }

        private void OnAssetSelected(FixedAssetLookupDto? asset)
        {
            selectedAsset = asset;
            if (asset != null)
            {
                model.AssetID = asset.AssetID;
                model.AssetDescription = asset.Description;
                if (asset.CalibrationFrequency.HasValue)
                {
                    model.CalibrationFrequency = asset.CalibrationFrequency.Value;
                    RecalculateNextDueDate();
                }
            }
            else
            {
                model.AssetID = string.Empty;
                model.AssetDescription = null;
            }
        }

        private Task<IEnumerable<VendorLookupDto>> SearchVendors(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<VendorLookupDto>>(vendors.Take(50));

            return Task.FromResult(vendors.Where(v =>
                v.AccNo.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                v.AccTitle.Contains(searchText, StringComparison.OrdinalIgnoreCase)
            ).Take(50));
        }

        private void OnVendorSelected(VendorLookupDto? vendor)
        {
            selectedVendor = vendor;
            if (vendor != null)
            {
                model.AccNo = vendor.AccNo;
                model.AccTitle = vendor.AccTitle;
            }
            else
            {
                model.AccNo = null;
                model.AccTitle = null;
            }
        }

        private void OnCalibrationDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out var dt))
            {
                model.DT = dt;
                RecalculateNextDueDate();
            }
        }

        private void OnNextDueDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out var dt))
            {
                model.NextDueDT = dt;
            }
            else
            {
                model.NextDueDT = null;
            }
        }

        private void OnFrequencyChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out var freq))
            {
                model.CalibrationFrequency = freq;
                RecalculateNextDueDate();
            }
        }

        private void RecalculateNextDueDate()
        {
            var baseDate = model.DT ?? DateTime.Today;
            if (model.CalibrationFrequency == 0) // Monthly
            {
                model.NextDueDT = baseDate.AddMonths(1);
            }
            else if (model.CalibrationFrequency == 1) // Yearly
            {
                model.NextDueDT = baseDate.AddYears(1);
            }
            else
            {
                model.NextDueDT = baseDate.AddDays(7);
            }
        }

        private void HandleAttachmentCountChanged(int count)
        {
            attachmentCount = count;
        }

        private async Task SaveCalibration(bool saveAndNew)
        {
            if (string.IsNullOrWhiteSpace(model.AssetID))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Please select a Fixed Asset.",
                    Duration = 3500
                });
                return;
            }

            if (!model.DT.HasValue)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Calibration Date is required.",
                    Duration = 3500
                });
                return;
            }

            if (model.NextDueDT.HasValue && model.NextDueDT.Value < model.DT.Value)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Invalid Next Schedule Date. Next Due Date cannot be earlier than Calibration Date.",
                    Duration = 4000
                });
                return;
            }

            isSaving = true;
            try
            {
                var userName = !string.IsNullOrWhiteSpace(PermissionService.CurrentUserName) ? PermissionService.CurrentUserName : "System";
                var savedId = await CalibrationData.SaveCalibrationAsync(model, userName, Environment.MachineName);
                model.EntryID = savedId;
                Id = savedId;

                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Calibration Saved",
                    Detail = $"Calibration for Asset {model.AssetID} saved successfully.",
                    Duration = 3500
                });

                if (saveAndNew)
                {
                    model = new CalibrationModel
                    {
                        DT = DateTime.Today,
                        NextDueDT = DateTime.Today.AddDays(7),
                        CalibrationFrequency = 0
                    };
                    selectedAsset = null;
                    selectedVendor = null;
                    Id = 0;
                    activeTab = "details";
                }
                else
                {
                    NavigationManager.NavigateTo(ReturnUrl ?? "/qms/calibrations");
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Save Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                isSaving = false;
            }
        }

        private void Cancel()
        {
            NavigationManager.NavigateTo(ReturnUrl ?? "/qms/calibrations");
        }
    }
}

