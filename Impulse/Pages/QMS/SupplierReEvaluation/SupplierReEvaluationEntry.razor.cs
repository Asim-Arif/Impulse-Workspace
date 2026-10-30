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
using Microsoft.AspNetCore.Components.Authorization;
using Radzen;

namespace Impulse.Pages.QMS.SupplierReEvaluation
{
    public partial class SupplierReEvaluationEntry : ComponentBase
    {
        [Parameter] public int Id { get; set; }
        [SupplyParameterFromQuery] public string? returnUrl { get; set; }

        [Inject] private ISupplierReEvaluationDataAccess ReEvaluationDataAccess { get; set; } = default!;
        [Inject] private ISupplierEvaluationDataAccess SupplierEvaluationDataAccess { get; set; } = default!;
        [Inject] private IEntityAttachmentService AttachmentService { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = default!;

        private SupplierReEvaluationModel model = new();
        private List<SupplierDropdownDto> suppliers = new();
        private SupplierDropdownDto? selectedSupplier;

        private bool isLoading = true;
        private bool isSaving = false;
        private bool isEditMode => Id > 0;
        private int attachmentCount = 0;

        private AttachmentModal? attachmentModalRef;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                isLoading = true;
                await LoadSuppliersAsync();

                if (Id > 0)
                {
                    var existing = await ReEvaluationDataAccess.GetReEvaluationByIdAsync(Id);
                    if (existing != null)
                    {
                        model = existing;
                        selectedSupplier = suppliers.FirstOrDefault(s => s.AccNo == model.AccNo);
                        await LoadAttachmentCountAsync();
                    }
                    else
                    {
                        NotificationService.Notify(NotificationSeverity.Warning, "Not Found", $"Re-Evaluation #{Id} was not found.");
                        NavigateBack();
                        return;
                    }
                }
                else
                {
                    var authState = await AuthStateProvider.GetAuthenticationStateAsync();
                    string userName = authState.User.Identity?.Name ?? Environment.UserName;

                    model = new SupplierReEvaluationModel
                    {
                        ReviewDT = DateTime.Today,
                        EntryDT = DateTime.Today,
                        PreparedBy = userName,
                        UserName = userName,
                        MachineName = Environment.MachineName,
                        // Default starting scores to 0 so they are ready for input
                        QualityOfProduction = 0,
                        Price = 0,
                        Packing = 0,
                        RejectionRate = 0,
                        DeliveryInTime = 0,
                        ReWorkRate = 0,
                        NoOfOrdersCompleted = 0,
                        CancelledOrders = 0,
                        ResponseTime = 0,
                        TechnicalAbility = 0,
                        ProductionCapacity = 0,
                        EnvironmentRequirements = 0,
                        LightingEtc = 0
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
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to load suppliers: {ex.Message}");
            }
        }

        private async Task LoadAttachmentCountAsync()
        {
            if (model.EntryID > 0)
            {
                try
                {
                    attachmentCount = await AttachmentService.GetCountAsync("SupplierReEvaluation", model.EntryID.ToString());
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

        private void OnSupplierSelected(SupplierDropdownDto? sup)
        {
            selectedSupplier = sup;
            if (sup != null)
            {
                model.AccNo = sup.AccNo;
                model.AccTitle = sup.AccTitle;
            }
            else
            {
                model.AccNo = string.Empty;
                model.AccTitle = string.Empty;
            }
        }

        private void OnDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out DateTime dt))
            {
                model.ReviewDT = dt;
            }
        }

        private async Task SaveReEvaluationAsync()
        {
            if (string.IsNullOrWhiteSpace(model.AccNo))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Please select a Supplier / Maker.");
                return;
            }

            if (!model.ReviewDT.HasValue)
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Please select a valid Review Date.");
                return;
            }

            try
            {
                isSaving = true;

                var authState = await AuthStateProvider.GetAuthenticationStateAsync();
                model.UserName ??= authState.User.Identity?.Name ?? Environment.UserName;
                model.MachineName ??= Environment.MachineName;

                int savedId = await ReEvaluationDataAccess.SaveReEvaluationAsync(model);
                if (savedId > 0)
                {
                    model.EntryID = savedId;
                    NotificationService.Notify(NotificationSeverity.Success, "Saved", $"Supplier Re-Evaluation #{savedId} saved successfully.");
                    NavigateBack();
                }
                else
                {
                    NotificationService.Notify(NotificationSeverity.Error, "Save Failed", "Could not save re-evaluation record.");
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
            string entityId = model.EntryID > 0 ? model.EntryID.ToString() : model.AccNo;
            if (string.IsNullOrWhiteSpace(entityId)) return;

            attachmentModalRef?.Open(
                entityType: "SupplierReEvaluation",
                entityId: entityId,
                secondaryId: model.AccNo,
                displayName: $"Supplier Re-Eval #{model.EntryID}: {model.AccTitle ?? model.AccNo} ({model.AccNo})"
            );
        }

        private async Task PrintCurrentReEvaluation()
        {
            if (model.EntryID <= 0) return;

            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = ReportNames.QMS.SupplierReEvaluation,
                    SelectionFormula = $"{{SupplierReEvaluation.EntryID}} = {model.EntryID}"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", $"Could not print evaluation slip: {ex.Message}");
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
                NavigationManager.NavigateTo("/qms/supplier-reevaluations");
            }
        }
    }
}

