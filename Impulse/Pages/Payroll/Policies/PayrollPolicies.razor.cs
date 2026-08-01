using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.Payroll.Policies
{
    public partial class PayrollPolicies : ComponentBase
    {
        [Inject] private IPayrollPoliciesService PoliciesService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;

        [Parameter, SupplyParameterFromQuery]
        public string? ReturnUrl { get; set; }

        private List<EmpAttCategoryModel> Categories { get; set; } = new();
        private EmpAttCategoryModel? SelectedCategory { get; set; }
        private EmpAttCategoryModel Model { get; set; } = new();

        private bool IsLoading { get; set; } = false;
        private bool IsSaving { get; set; } = false;
        private bool IsDeleting { get; set; } = false;

        // Add Category Modal State
        private bool ShowAddModal { get; set; } = false;
        private EmpAttCategoryModel NewCategoryModel { get; set; } = new();

        protected override async Task OnInitializedAsync()
        {
            await LoadCategoriesAsync();
        }

        private async Task LoadCategoriesAsync()
        {
            IsLoading = true;
            StateHasChanged();
            try
            {
                Categories = await PoliciesService.GetCategoriesAsync();
                if (Categories.Any())
                {
                    SelectedCategory = Categories.First();
                    SelectCategoryModel(SelectedCategory);
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Policies",
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

        private void SelectCategoryModel(EmpAttCategoryModel cat)
        {
            Model = new EmpAttCategoryModel
            {
                EntryID = cat.EntryID,
                LunchTimeScanning = cat.LunchTimeScanning,
                LateDeduction = cat.LateDeduction,
                LateDeductionAfter = cat.LateDeductionAfter,
                GatePassRequired = cat.GatePassRequired,
                LeavesRecorded = cat.LeavesRecorded,
                OverTimePaid = cat.OverTimePaid,
                DeductionSpecialAbsents = cat.DeductionSpecialAbsents,
                OTRate = cat.OTRate,
                DeductionPercentage = cat.DeductionPercentage,
                Remarks = cat.Remarks,
                Saturday_Early_Out_Mins = cat.Saturday_Early_Out_Mins
            };
        }

        private void OnCategoryChanged(ChangeEventArgs e)
        {
            if (int.TryParse(Convert.ToString(e.Value), out int catId))
            {
                SelectedCategory = Categories.FirstOrDefault(c => c.EntryID == catId);
                if (SelectedCategory != null)
                {
                    SelectCategoryModel(SelectedCategory);
                }
            }
        }

        private void OpenAddModal()
        {
            NewCategoryModel = new EmpAttCategoryModel
            {
                Remarks = $"Category #{Categories.Count + 1}",
                LunchTimeScanning = false,
                LateDeduction = false,
                LateDeductionAfter = "10",
                GatePassRequired = false,
                LeavesRecorded = false,
                OverTimePaid = false,
                DeductionSpecialAbsents = false,
                OTRate = 1.5f,
                DeductionPercentage = 100,
                Saturday_Early_Out_Mins = 0
            };
            ShowAddModal = true;
        }

        private void CloseAddModal()
        {
            ShowAddModal = false;
            NewCategoryModel = new EmpAttCategoryModel();
        }

        private async Task SaveCategoryAsync(bool isNewFromModal)
        {
            var targetModel = isNewFromModal ? NewCategoryModel : Model;

            if (string.IsNullOrWhiteSpace(targetModel.Remarks))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Please enter a Category Title / Remarks.",
                    Duration = 3000
                });
                return;
            }

            IsSaving = true;
            StateHasChanged();
            try
            {
                bool success = await PoliciesService.SaveCategoryAsync(targetModel);
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Category Saved",
                        Detail = $"Attendance category '{targetModel.Remarks}' saved successfully.",
                        Duration = 4000
                    });

                    if (isNewFromModal)
                    {
                        CloseAddModal();
                    }

                    int savedId = targetModel.EntryID;
                    await LoadCategoriesAsync();
                    SelectedCategory = Categories.FirstOrDefault(c => c.EntryID == savedId) ?? Categories.FirstOrDefault();
                    if (SelectedCategory != null)
                    {
                        SelectCategoryModel(SelectedCategory);
                    }
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

        private async Task DeleteCategoryAsync()
        {
            if (SelectedCategory == null) return;

            IsDeleting = true;
            StateHasChanged();
            try
            {
                bool success = await PoliciesService.DeleteCategoryAsync(SelectedCategory.EntryID);
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Category Deleted",
                        Detail = $"Category #{SelectedCategory.EntryID} deleted.",
                        Duration = 3000
                    });
                    await LoadCategoriesAsync();
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
