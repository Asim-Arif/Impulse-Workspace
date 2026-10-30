using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using BlazorContextMenu;
using DataAccessLibrary.Interface.QMS;
using DataAccessLibrary.Models.QMS;
using Impulse.Components.Common;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;

namespace Impulse.Pages.QMS.Documents
{
    public partial class QMSDocumentsList : ComponentBase
    {
        [Inject] private IQMSDocumentDataAccess DocumentDataAccess { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigationService { get; set; } = default!;

        private List<QMSDocumentListItemModel> documents = new();
        private List<string> documentTypes = new();
        private AttachmentModal? attachmentModalRef;
        private QMSDocumentListItemModel? documentToDelete;

        private bool isLoading = true;
        private string searchText = string.Empty;
        private string selectedDocType = string.Empty;
        private int selectedInternalFilter = -1; // -1: All, 1: Internal, 0: External
        private string selectedPeriodPreset = "All";
        private DateTime? filterDateFrom;
        private DateTime? filterDateTo;

        private int selectedCount => documents.Count(d => d.IsSelected);
        private bool isAllSelected => documents.Count > 0 && documents.All(d => d.IsSelected);

        protected override async Task OnInitializedAsync()
        {
            await LoadDocumentTypes();
            ApplyPeriodPreset("All");
            await LoadDocuments();
        }

        private async Task LoadDocumentTypes()
        {
            try
            {
                documentTypes = await DocumentDataAccess.GetDocumentTypesListAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to load document types: {ex.Message}");
            }
        }

        private async Task LoadDocuments()
        {
            try
            {
                isLoading = true;
                var filter = new QMSDocumentFilterDto
                {
                    SearchText = searchText,
                    DocumentType = selectedDocType,
                    InternalFilter = selectedInternalFilter >= 0 ? selectedInternalFilter : null,
                    DateFrom = filterDateFrom,
                    DateTo = filterDateTo
                };

                documents = await DocumentDataAccess.GetDocumentsAsync(filter);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to load documents: {ex.Message}");
            }
            finally
            {
                isLoading = false;
            }
        }

        private void OnSearchInput(ChangeEventArgs e)
        {
            searchText = e.Value?.ToString() ?? string.Empty;
            _ = LoadDocuments();
        }

        private void OnDocTypeChanged(ChangeEventArgs e)
        {
            selectedDocType = e.Value?.ToString() ?? string.Empty;
            _ = LoadDocuments();
        }

        private void OnScopeFilterChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int val))
            {
                selectedInternalFilter = val;
            }
            else
            {
                selectedInternalFilter = -1;
            }
            _ = LoadDocuments();
        }

        private void OnPeriodPresetChanged(ChangeEventArgs e)
        {
            selectedPeriodPreset = e.Value?.ToString() ?? "All";
            ApplyPeriodPreset(selectedPeriodPreset);
            _ = LoadDocuments();
        }

        private void ApplyPeriodPreset(string preset)
        {
            var today = DateTime.Today;
            switch (preset)
            {
                case "ThisMonth":
                    filterDateFrom = new DateTime(today.Year, today.Month, 1);
                    filterDateTo = filterDateFrom.Value.AddMonths(1).AddDays(-1);
                    break;
                case "Last30Days":
                    filterDateFrom = today.AddDays(-30);
                    filterDateTo = today;
                    break;
                case "ThisYear":
                    filterDateFrom = new DateTime(today.Year, 1, 1);
                    filterDateTo = new DateTime(today.Year, 12, 31);
                    break;
                case "All":
                default:
                    filterDateFrom = null;
                    filterDateTo = null;
                    break;
            }
        }

        private async Task ResetFilters()
        {
            searchText = string.Empty;
            selectedDocType = string.Empty;
            selectedInternalFilter = -1;
            selectedPeriodPreset = "All";
            filterDateFrom = null;
            filterDateTo = null;
            await LoadDocuments();
        }

        private async Task RefreshData()
        {
            await LoadDocuments();
        }

        private void ToggleSelectAll(ChangeEventArgs e)
        {
            bool check = (bool)(e.Value ?? false);
            foreach (var doc in documents)
            {
                doc.IsSelected = check;
            }
        }

        private void ToggleRowSelection(QMSDocumentListItemModel doc, ChangeEventArgs e)
        {
            doc.IsSelected = (bool)(e.Value ?? false);
        }

        private void NavigateToHub()
        {
            NavigationManager.NavigateTo("/qms");
        }

        private void NavigateToCreate()
        {
            NavigationManager.NavigateTo("/qms/documents/new?returnUrl=/qms/documents");
        }

        private void NavigateToEdit(int entryId)
        {
            NavigationManager.NavigateTo($"/qms/documents/{entryId}?returnUrl=/qms/documents");
        }

        private void OpenAttachmentModal(int entryId, string title)
        {
            attachmentModalRef?.Open(
                entityType: "QMSDocument",
                entityId: entryId.ToString(),
                secondaryId: null,
                displayName: $"Document: {title}"
            );
        }

        private void PromptDeleteDocument(QMSDocumentListItemModel doc)
        {
            documentToDelete = doc;
        }

        private void CancelDelete()
        {
            documentToDelete = null;
        }

        private async Task ConfirmDelete()
        {
            if (documentToDelete == null) return;

            try
            {
                bool success = await DocumentDataAccess.DeleteDocumentAsync(documentToDelete.EntryID);
                if (success)
                {
                    NotificationService.Notify(NotificationSeverity.Success, "Deleted", "Document deleted successfully.");
                    documents.Remove(documentToDelete);
                }
                else
                {
                    NotificationService.Notify(NotificationSeverity.Warning, "Warning", "Could not delete document.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to delete document: {ex.Message}");
            }
            finally
            {
                documentToDelete = null;
            }
        }

        private async Task PrintSelectedDistributionList()
        {
            var selectedIds = documents.Where(d => d.IsSelected).Select(d => d.EntryID).ToList();
            if (!selectedIds.Any())
            {
                NotificationService.Notify(NotificationSeverity.Info, "Information", "No documents selected.");
                return;
            }

            string idList = string.Join(",", selectedIds);
            string selectionFormula = $"{{QMS_Documents.EntryID}} IN [{idList}]";
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "QMSDocument_DistributionList.rpt",
                SelectionFormula = selectionFormula
            });
        }

        private void OnContextMenuEdit(QMSDocumentListItemModel? doc)
        {
            if (doc != null)
            {
                NavigateToEdit(doc.EntryID);
            }
        }

        private async Task OnContextMenuPrintDist(QMSDocumentListItemModel? doc)
        {
            if (doc != null)
            {
                string selectionFormula = $"{{QMS_Documents.EntryID}} = {doc.EntryID}";
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "QMSDocument_DistributionList.rpt",
                    SelectionFormula = selectionFormula
                });
            }
        }

        private void OnContextMenuAttachments(QMSDocumentListItemModel? doc)
        {
            if (doc != null)
            {
                OpenAttachmentModal(doc.EntryID, doc.DocumentName ?? doc.DocumentNumber ?? "");
            }
        }

        private void OnContextMenuDelete(QMSDocumentListItemModel? doc)
        {
            if (doc != null)
            {
                PromptDeleteDocument(doc);
            }
        }
    }
}

