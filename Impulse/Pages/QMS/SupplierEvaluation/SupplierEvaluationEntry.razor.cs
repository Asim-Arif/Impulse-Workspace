using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.QMS;
using DataAccessLibrary.Models.QMS;
using Impulse.Components.Common;
using Impulse.Constants;
using Impulse.Services;
using Impulse.Services.Common;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.QMS.SupplierEvaluation
{
    public partial class SupplierEvaluationEntry : ComponentBase
    {
        [Parameter] public string? AccNo { get; set; }
        [SupplyParameterFromQuery] public string? returnUrl { get; set; }

        [Inject] private ISupplierEvaluationDataAccess SupplierEvaluationDataAccess { get; set; } = default!;
        [Inject] private IEntityAttachmentService AttachmentService { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;

        private SupplierEvaluationModel model = new();
        private List<SupplierDropdownDto> suppliers = new();
        private SupplierDropdownDto? selectedSupplier;

        private bool isLoading = true;
        private bool isSaving = false;
        private bool isEditMode => !string.IsNullOrWhiteSpace(AccNo);
        private int attachmentCount = 0;

        private AttachmentModal? attachmentModalRef;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                isLoading = true;
                await LoadSuppliersAsync();

                if (!string.IsNullOrWhiteSpace(AccNo))
                {
                    string decodedAccNo = Uri.UnescapeDataString(AccNo);
                    var existing = await SupplierEvaluationDataAccess.GetEvaluationByAccNoAsync(decodedAccNo);
                    if (existing != null)
                    {
                        model = existing;
                        selectedSupplier = suppliers.FirstOrDefault(s => s.AccNo == model.AccNo);
                        await LoadAttachmentCountAsync();
                    }
                    else
                    {
                        // Supplier exists, but evaluation not created yet
                        selectedSupplier = suppliers.FirstOrDefault(s => s.AccNo == decodedAccNo);
                        model = new SupplierEvaluationModel
                        {
                            AccNo = decodedAccNo,
                            AccTitle = selectedSupplier?.AccTitle,
                            EntryDT = DateTime.Today,
                            MarketingReputation = -1,
                            TechnicalCapability = -1
                        };
                    }
                }
                else
                {
                    model = new SupplierEvaluationModel
                    {
                        EntryDT = DateTime.Today,
                        MarketingReputation = -1,
                        TechnicalCapability = -1
                    };
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Initialization Error", $"Failed to load data: {ex.Message}");
            }
            finally
            {
                isLoading = false;
            }
        }

        private async Task LoadSuppliersAsync()
        {
            try
            {
                suppliers = await SupplierEvaluationDataAccess.GetSuppliersListAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to load suppliers list: {ex.Message}");
            }
        }

        private async Task LoadAttachmentCountAsync()
        {
            if (!string.IsNullOrWhiteSpace(model.AccNo))
            {
                try
                {
                    attachmentCount = await AttachmentService.GetCountAsync("SupplierEvaluation", model.AccNo);
                }
                catch
                {
                    attachmentCount = 0;
                }
            }
        }

        private Task<IEnumerable<SupplierDropdownDto>> SearchSuppliers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return Task.FromResult<IEnumerable<SupplierDropdownDto>>(suppliers.Take(100));
            }

            var query = searchText.Trim();
            var matches = suppliers.Where(s =>
                (!string.IsNullOrEmpty(s.AccNo) && s.AccNo.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(s.AccTitle) && s.AccTitle.Contains(query, StringComparison.OrdinalIgnoreCase))
            );

            return Task.FromResult<IEnumerable<SupplierDropdownDto>>(matches);
        }

        private async Task OnSupplierSelected(SupplierDropdownDto? sup)
        {
            selectedSupplier = sup;
            if (sup == null)
            {
                model = new SupplierEvaluationModel
                {
                    EntryDT = DateTime.Today,
                    MarketingReputation = -1,
                    TechnicalCapability = -1
                };
                attachmentCount = 0;
                return;
            }

            try
            {
                // Check if an evaluation already exists for this supplier (mirroring legacy AddEvaluation logic)
                var existing = await SupplierEvaluationDataAccess.GetEvaluationByAccNoAsync(sup.AccNo);
                if (existing != null)
                {
                    model = existing;
                    NotificationService.Notify(NotificationSeverity.Info, "Existing Record Loaded", $"Found existing evaluation for {sup.AccTitle}.");
                }
                else
                {
                    model = new SupplierEvaluationModel
                    {
                        AccNo = sup.AccNo,
                        AccTitle = sup.AccTitle,
                        EntryDT = DateTime.Today,
                        MarketingReputation = -1,
                        TechnicalCapability = -1
                    };
                }

                await LoadAttachmentCountAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Could not check existing evaluation: {ex.Message}");
            }
        }

        private void OnDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out DateTime dt))
            {
                model.EntryDT = dt;
            }
        }

        private async Task SaveEvaluationAsync()
        {
            if (string.IsNullOrWhiteSpace(model.AccNo))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Please select a Supplier / Maker.");
                return;
            }

            if (!model.MarketingReputation.HasValue || model.MarketingReputation.Value < 0)
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Please select Marketing Reputation (Poor, Avg, or Good).");
                return;
            }

            if (!model.TechnicalCapability.HasValue || model.TechnicalCapability.Value < 0)
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Please select Technical Capability (Poor, Avg, or Good).");
                return;
            }

            try
            {
                isSaving = true;
                bool success = await SupplierEvaluationDataAccess.SaveEvaluationAsync(model);
                if (success)
                {
                    NotificationService.Notify(NotificationSeverity.Success, "Saved", "Supplier evaluation saved successfully.");
                    NavigateBack();
                }
                else
                {
                    NotificationService.Notify(NotificationSeverity.Error, "Save Failed", "Could not save evaluation record.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Save operation failed: {ex.Message}");
            }
            finally
            {
                isSaving = false;
            }
        }

        private void OpenAttachmentModal()
        {
            if (string.IsNullOrWhiteSpace(model.AccNo)) return;

            attachmentModalRef?.Open(
                entityType: "SupplierEvaluation",
                entityId: model.AccNo,
                secondaryId: model.EntryID.ToString(),
                displayName: $"Supplier: {model.AccTitle ?? model.AccNo} ({model.AccNo})"
            );
        }

        private async Task PrintCurrentEvaluation()
        {
            if (string.IsNullOrWhiteSpace(model.AccNo)) return;

            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = ReportNames.QMS.SupplierEvaluation,
                    SelectionFormula = $"{{SupplierEvaluation.AccNo}} = '{model.AccNo}'"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", $"Could not print supplier evaluation: {ex.Message}");
            }
        }

        private void NavigateBack()
        {
            if (!string.IsNullOrWhiteSpace(returnUrl))
            {
                NavigationManager.NavigateTo(returnUrl);
            }
            else
            {
                NavigationManager.NavigateTo("/qms/supplier-evaluations");
            }
        }
    }
}

