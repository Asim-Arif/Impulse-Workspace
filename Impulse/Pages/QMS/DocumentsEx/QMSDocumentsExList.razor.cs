using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.QMS;
using DataAccessLibrary.Models.QMS;
using Impulse.Components.Common;
using Impulse.Constants;
using Impulse.Services;
using Impulse.Services.Setup;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;

namespace Impulse.Pages.QMS.DocumentsEx
{
    public partial class QMSDocumentsExList : ComponentBase
    {
        [Parameter] public string? Category { get; set; }

        [Inject] private IQMSDocumentExDataAccess DocumentExData { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private NotificationService NotificationService { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private IUserPermissionService PermissionService { get; set; } = null!;

        private List<QMSDocumentExListItemDto> documents = new();
        private string searchText = string.Empty;
        private DateTime? fromDate;
        private DateTime? toDate;
        private bool isLoading = true;

        private AttachmentModal? attachmentModalRef;
        private QMSDocumentExListItemDto? itemToDelete;
        private bool showDeleteModal = false;

        private string MasterCategoryName = "Standards";

        private string CategoryTitle => MasterCategoryName switch
        {
            "Manuals" => "Manuals",
            "Procedures" => "Procedures",
            "Work Instructions" => "Work Instructions",
            "Technical Files" => "Technical Files",
            "Lean Management" => "Lean Management",
            _ => "Standards"
        };

        private string CategorySubtitle => MasterCategoryName switch
        {
            "Manuals" => "Quality Manuals & Organizational Policies Master",
            "Procedures" => "Standard Operating Procedures (SOPs) & Controlled Policies",
            "Work Instructions" => "Work Instructions & Operational Step-by-Step Guidelines",
            "Technical Files" => "Device & Medical Product Technical Master Files",
            "Lean Management" => "Continuous Improvement, 5S & Lean Documentation",
            _ => "QMS International, Regulatory & Quality Standards Master"
        };

        private string CategoryIcon => MasterCategoryName switch
        {
            "Manuals" => "fas fa-book",
            "Procedures" => "fas fa-tasks",
            "Work Instructions" => "fas fa-file-lines",
            "Technical Files" => "fas fa-file-shield",
            "Lean Management" => "fas fa-chart-line",
            _ => "fas fa-book-bookmark"
        };

        private string CategorySlug => MasterCategoryName switch
        {
            "Manuals" => "manuals",
            "Procedures" => "procedures",
            "Work Instructions" => "work-instructions",
            "Technical Files" => "technical-files",
            "Lean Management" => "lean-management",
            _ => "standards"
        };

        private bool IsAllSelected => documents.Any() && documents.All(d => d.IsSelected);
        private int SelectedCount => documents.Count(d => d.IsSelected);

        protected override async Task OnInitializedAsync()
        {
            await PermissionService.InitializeAsync();
        }

        protected override async Task OnParametersSetAsync()
        {
            ResolveCategory();
            await LoadDocumentsAsync();
        }

        private void ResolveCategory()
        {
            var uri = NavigationManager.Uri.ToLowerInvariant();

            if (!string.IsNullOrWhiteSpace(Category))
            {
                var cat = Category.Trim().ToLowerInvariant();
                if (cat.Contains("manual")) MasterCategoryName = "Manuals";
                else if (cat.Contains("proc")) MasterCategoryName = "Procedures";
                else if (cat.Contains("work") || cat.Contains("instruction")) MasterCategoryName = "Work Instructions";
                else if (cat.Contains("tech") || cat.Contains("file")) MasterCategoryName = "Technical Files";
                else if (cat.Contains("lean")) MasterCategoryName = "Lean Management";
                else MasterCategoryName = "Standards";
                return;
            }

            if (uri.Contains("/qms/manuals")) MasterCategoryName = "Manuals";
            else if (uri.Contains("/qms/procedures")) MasterCategoryName = "Procedures";
            else if (uri.Contains("/qms/work-instructions")) MasterCategoryName = "Work Instructions";
            else if (uri.Contains("/qms/technical-files")) MasterCategoryName = "Technical Files";
            else if (uri.Contains("/qms/lean-management")) MasterCategoryName = "Lean Management";
            else MasterCategoryName = "Standards";
        }

        private async Task LoadDocumentsAsync()
        {
            isLoading = true;
            try
            {
                documents = await DocumentExData.GetDocumentsAsync(MasterCategoryName, searchText, fromDate, toDate);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error Loading Documents", ex.Message);
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
                await LoadDocumentsAsync();
            }
        }

        private async Task ResetFilters()
        {
            searchText = string.Empty;
            fromDate = null;
            toDate = null;
            await LoadDocumentsAsync();
        }

        private void ToggleSelectAll(ChangeEventArgs e)
        {
            bool selectAll = (bool)(e.Value ?? false);
            foreach (var doc in documents)
            {
                doc.IsSelected = selectAll;
            }
        }

        private void NavigateToCreate()
        {
            var returnUrl = Uri.EscapeDataString(NavigationManager.Uri);
            NavigationManager.NavigateTo($"/qms/documents-ex/entry/{CategorySlug}?returnUrl={returnUrl}");
        }

        private void NavigateToEdit(int entryId)
        {
            var returnUrl = Uri.EscapeDataString(NavigationManager.Uri);
            NavigationManager.NavigateTo($"/qms/documents-ex/entry/{CategorySlug}/{entryId}?returnUrl={returnUrl}");
        }

        private void OpenAttachments(QMSDocumentExListItemDto item)
        {
            attachmentModalRef?.Open(item.EntryID.ToString(), "QMSDocumentEx", item.DocumentName ?? item.DocumentNumber ?? "Document");
        }

        private async Task PrintSelectedDistributionList()
        {
            var selectedIds = documents.Where(d => d.IsSelected).Select(d => d.EntryID).ToList();
            if (!selectedIds.Any())
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Selection Required", "Please select at least one document to print distribution list.");
                return;
            }

            try
            {
                string idList = string.Join(",", selectedIds);
                string selectionFormula = $"{{QMS_Documents_Ex.EntryID}} in [{idList}]";

                var request = new ReportRequest
                {
                    ReportName = ReportNames.QMS.QMSDocumentDistributionList,
                    SelectionFormula = selectionFormula
                };

                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", $"Could not print distribution list: {ex.Message}");
            }
        }

        private async Task PrintDistributionListSingle(QMSDocumentExListItemDto item)
        {
            try
            {
                string selectionFormula = $"{{QMS_Documents_Ex.EntryID}} = {item.EntryID}";
                var request = new ReportRequest
                {
                    ReportName = ReportNames.QMS.QMSDocumentDistributionList,
                    SelectionFormula = selectionFormula
                };

                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", $"Could not print distribution list: {ex.Message}");
            }
        }

        private void ConfirmDelete(QMSDocumentExListItemDto item)
        {
            itemToDelete = item;
            showDeleteModal = true;
        }

        private void CloseDeleteModal()
        {
            showDeleteModal = false;
            itemToDelete = null;
        }

        private async Task ExecuteDelete()
        {
            if (itemToDelete == null) return;

            try
            {
                var success = await DocumentExData.DeleteDocumentAsync(itemToDelete.EntryID);
                if (success)
                {
                    documents.Remove(itemToDelete);
                    NotificationService.Notify(NotificationSeverity.Success, "Deleted", $"Document '{itemToDelete.DocumentNumber}' removed successfully.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Delete Failed", ex.Message);
            }
            finally
            {
                CloseDeleteModal();
            }
        }
    }
}

