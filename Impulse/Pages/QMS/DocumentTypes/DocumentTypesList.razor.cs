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

namespace Impulse.Pages.QMS.DocumentTypes
{
    public partial class DocumentTypesList : ComponentBase
    {
        [Inject] private IQMSDocumentTypeDataAccess DocumentTypeData { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private NotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IUserPermissionService PermissionService { get; set; } = null!;

        private List<QMSDocumentTypeModel> documentTypes = new();
        private string newDocumentTypeName = string.Empty;
        private string searchTerm = string.Empty;
        private bool isLoading = true;
        private bool isAdding = false;
        private bool isSavingEdit = false;

        private QMSDocumentTypeModel? itemToDelete;
        private bool showDeleteConfirmModal = false;

        protected override async Task OnInitializedAsync()
        {
            isLoading = true;
            try
            {
                await PermissionService.InitializeAsync();
                await LoadDocumentTypesAsync();
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

        private async Task LoadDocumentTypesAsync()
        {
            isLoading = true;
            try
            {
                documentTypes = await DocumentTypeData.GetDocumentTypesAsync(searchTerm);
            }
            catch (Exception ex)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Document Types",
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
                await LoadDocumentTypesAsync();
            }
        }

        private async Task HandleNewTypeKeyUp(KeyboardEventArgs e)
        {
            if (e.Key == "Enter")
            {
                await AddDocumentType();
            }
        }

        private async Task AddDocumentType()
        {
            if (string.IsNullOrWhiteSpace(newDocumentTypeName))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Document Type name cannot be empty.",
                    Duration = 3500
                });
                return;
            }

            var trimmedName = newDocumentTypeName.Trim();

            isAdding = true;
            try
            {
                bool exists = await DocumentTypeData.DocumentTypeExistsAsync(trimmedName);
                if (exists)
                {
                    NotificationServiceManager.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Warning,
                        Summary = "Duplicate Entry",
                        Detail = $"Document Type '{trimmedName}' already exists.",
                        Duration = 4000
                    });
                    return;
                }

                await DocumentTypeData.AddDocumentTypeAsync(trimmedName);
                newDocumentTypeName = string.Empty;

                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Added Successfully",
                    Detail = $"Document Type '{trimmedName}' created.",
                    Duration = 3500
                });

                await LoadDocumentTypesAsync();
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

        private void StartEdit(QMSDocumentTypeModel item)
        {
            // Cancel any other item being edited
            foreach (var doc in documentTypes.Where(d => d.IsEditing && d.EntryID != item.EntryID))
            {
                doc.IsEditing = false;
            }

            item.IsEditing = true;
            item.EditBuffer = item.DocumentType;
        }

        private void CancelEdit(QMSDocumentTypeModel item)
        {
            item.IsEditing = false;
            item.EditBuffer = string.Empty;
        }

        private async Task HandleEditKeyUp(KeyboardEventArgs e, QMSDocumentTypeModel item)
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

        private async Task SaveEdit(QMSDocumentTypeModel item)
        {
            if (string.IsNullOrWhiteSpace(item.EditBuffer))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Document Type name cannot be empty.",
                    Duration = 3500
                });
                return;
            }

            var trimmed = item.EditBuffer.Trim();

            isSavingEdit = true;
            try
            {
                bool exists = await DocumentTypeData.DocumentTypeExistsAsync(trimmed, item.EntryID);
                if (exists)
                {
                    NotificationServiceManager.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Warning,
                        Summary = "Duplicate Entry",
                        Detail = $"Another Document Type with name '{trimmed}' already exists.",
                        Duration = 4000
                    });
                    return;
                }

                await DocumentTypeData.UpdateDocumentTypeAsync(item.EntryID, trimmed);
                item.DocumentType = trimmed;
                item.IsEditing = false;

                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Updated",
                    Detail = $"Document Type updated to '{trimmed}'.",
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

        private void ConfirmDelete(QMSDocumentTypeModel item)
        {
            if (item.UsageCount > 0)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Cannot Delete",
                    Detail = $"This Document Type cannot be deleted because it is assigned to {item.UsageCount} active document(s) in Document Control.",
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
                var success = await DocumentTypeData.DeleteDocumentTypeAsync(itemToDelete.EntryID);
                if (success)
                {
                    documentTypes.Remove(itemToDelete);
                    NotificationServiceManager.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Deleted",
                        Detail = $"Document Type '{itemToDelete.DocumentType}' removed.",
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

