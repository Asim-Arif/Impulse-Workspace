using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.QMS;
using DataAccessLibrary.Models.QMS;
using Impulse.Services.Setup;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;

namespace Impulse.Pages.QMS.Reasons
{
    public partial class QMSReasonsList : ComponentBase
    {
        [Inject] private IQMSReasonDataAccess ReasonDataAccess { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private NotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IUserPermissionService PermissionService { get; set; } = null!;

        private List<QMSReasonModel> reasons = new();
        private string newReasonText = string.Empty;
        private string searchTerm = string.Empty;
        private bool isLoading = true;
        private bool isAdding = false;
        private bool isSavingEdit = false;

        private QMSReasonModel? itemToDelete;
        private bool showDeleteConfirmModal = false;

        protected override async Task OnInitializedAsync()
        {
            isLoading = true;
            try
            {
                await PermissionService.InitializeAsync();
                await LoadReasonsAsync();
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

        private async Task LoadReasonsAsync()
        {
            isLoading = true;
            try
            {
                reasons = await ReasonDataAccess.GetAllReasonsAsync(searchTerm);
            }
            catch (Exception ex)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Reasons",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                isLoading = false;
            }
        }

        private async Task HandleSearchKeyUp(KeyboardEventArgs e)
        {
            if (e.Key == "Enter")
            {
                await LoadReasonsAsync();
            }
        }

        private async Task HandleNewReasonKeyUp(KeyboardEventArgs e)
        {
            if (e.Key == "Enter")
            {
                await AddReason();
            }
        }

        private async Task AddReason()
        {
            if (string.IsNullOrWhiteSpace(newReasonText))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Reason Description cannot be empty.",
                    Duration = 3500
                });
                return;
            }

            var trimmedReason = newReasonText.Trim();

            if (trimmedReason.Length > 50)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Reason Description cannot exceed 50 characters.",
                    Duration = 3500
                });
                return;
            }

            isAdding = true;
            try
            {
                bool exists = await ReasonDataAccess.ReasonExistsAsync(trimmedReason);
                if (exists)
                {
                    NotificationServiceManager.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Warning,
                        Summary = "Duplicate Entry",
                        Detail = $"Reason '{trimmedReason}' already exists.",
                        Duration = 4000
                    });
                    return;
                }

                await ReasonDataAccess.AddReasonAsync(trimmedReason);
                newReasonText = string.Empty;

                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Added Successfully",
                    Detail = $"Reason '{trimmedReason}' created.",
                    Duration = 3500
                });

                await LoadReasonsAsync();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Add Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                isAdding = false;
            }
        }

        private void StartEdit(QMSReasonModel item)
        {
            // Cancel any other item being edited
            foreach (var r in reasons.Where(x => x.IsEditing && x.EntryID != item.EntryID))
            {
                r.IsEditing = false;
            }

            item.IsEditing = true;
            item.EditBuffer = item.Reason_QMS;
        }

        private void CancelEdit(QMSReasonModel item)
        {
            item.IsEditing = false;
            item.EditBuffer = string.Empty;
        }

        private async Task HandleEditKeyUp(KeyboardEventArgs e, QMSReasonModel item)
        {
            if (e.Key == "Enter")
            {
                await SaveEdit(item);
            }
            else if (e.Key == "Escape")
            {
                CancelEdit(item);
            }
        }

        private async Task SaveEdit(QMSReasonModel item)
        {
            if (string.IsNullOrWhiteSpace(item.EditBuffer))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Reason Description cannot be empty.",
                    Duration = 3500
                });
                return;
            }

            var trimmed = item.EditBuffer.Trim();

            if (trimmed.Length > 50)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Reason Description cannot exceed 50 characters.",
                    Duration = 3500
                });
                return;
            }

            if (string.Equals(trimmed, item.Reason_QMS, StringComparison.OrdinalIgnoreCase))
            {
                item.IsEditing = false;
                return;
            }

            isSavingEdit = true;
            try
            {
                bool exists = await ReasonDataAccess.ReasonExistsAsync(trimmed, item.EntryID);
                if (exists)
                {
                    NotificationServiceManager.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Warning,
                        Summary = "Duplicate Entry",
                        Detail = $"Another Reason with description '{trimmed}' already exists.",
                        Duration = 4000
                    });
                    return;
                }

                await ReasonDataAccess.UpdateReasonAsync(item.EntryID, item.Reason_QMS, trimmed);
                item.Reason_QMS = trimmed;
                item.IsEditing = false;

                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Updated",
                    Detail = $"Reason updated to '{trimmed}'.",
                    Duration = 3500
                });
            }
            catch (Exception ex)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Update Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                isSavingEdit = false;
            }
        }

        private void ConfirmDelete(QMSReasonModel item)
        {
            if (item.UsageCount > 0)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Cannot Delete",
                    Detail = $"This Reason cannot be deleted because it is associated with {item.UsageCount} complaint item record(s).",
                    Duration = 5000
                });
                return;
            }

            itemToDelete = item;
            showDeleteConfirmModal = true;
        }

        private void CloseDeleteModal()
        {
            showDeleteConfirmModal = false;
            itemToDelete = null;
        }

        private async Task ExecuteDelete()
        {
            if (itemToDelete == null) return;

            try
            {
                var success = await ReasonDataAccess.DeleteReasonAsync(itemToDelete.EntryID, itemToDelete.Reason_QMS);
                if (success)
                {
                    reasons.Remove(itemToDelete);
                    NotificationServiceManager.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Deleted",
                        Detail = $"Reason '{itemToDelete.Reason_QMS}' removed.",
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
            finally
            {
                CloseDeleteModal();
            }
        }
    }
}

