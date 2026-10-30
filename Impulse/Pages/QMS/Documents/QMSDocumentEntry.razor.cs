using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.QMS;
using DataAccessLibrary.Models.QMS;
using Impulse.Services.Common;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Radzen;

namespace Impulse.Pages.QMS.Documents
{
    public partial class QMSDocumentEntry : ComponentBase
    {
        [Parameter] public int Id { get; set; }
        [SupplyParameterFromQuery] public string? returnUrl { get; set; }

        [Inject] private IQMSDocumentDataAccess DocumentDataAccess { get; set; } = default!;
        [Inject] private IEntityAttachmentService AttachmentService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = default!;

        private QMSDocumentModel model = new();
        private List<EmployeeDropdownDto> employees = new();
        private List<string> documentTypes = new();
        private List<int> deletedDistributionIds = new();

        private EmployeeDropdownDto? selectedInitiator;
        private EmployeeDropdownDto? selectedDistEmployee;

        private string activeTab = "details";
        private bool isLoading = true;
        private bool isSaving = false;
        private int attachmentCount = 0;

        private int activeDistributionCount => model.DistributionList.Count(d => !d.IsMarkedForDeletion);

        protected override async Task OnInitializedAsync()
        {
            try
            {
                isLoading = true;
                await Task.WhenAll(LoadEmployees(), LoadDocumentTypes());

                if (Id > 0)
                {
                    var loaded = await DocumentDataAccess.GetDocumentByIdAsync(Id);
                    if (loaded != null)
                    {
                        model = loaded;
                        if (!string.IsNullOrWhiteSpace(model.InitiaterEmpID))
                        {
                            selectedInitiator = employees.FirstOrDefault(e => e.EmpID == model.InitiaterEmpID);
                            if (selectedInitiator != null)
                            {
                                model.InitiatorDepartment = selectedInitiator.DeptName;
                                model.InitiatorDesignation = selectedInitiator.Designation;
                            }
                        }
                    }
                    else
                    {
                        NotificationService.Notify(NotificationSeverity.Warning, "Not Found", $"Document #{Id} was not found.");
                        NavigateBack();
                        return;
                    }

                    await LoadAttachmentCount();
                }
                else
                {
                    model = new QMSDocumentModel
                    {
                        IssueDT = DateTime.Today,
                        Revision = "00",
                        Internal = true
                    };
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to initialize form: {ex.Message}");
            }
            finally
            {
                isLoading = false;
            }
        }

        private async Task LoadEmployees()
        {
            try
            {
                employees = await DocumentDataAccess.GetEmployeesAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to load employees: {ex.Message}");
            }
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

        private async Task LoadAttachmentCount()
        {
            if (Id > 0)
            {
                try
                {
                    attachmentCount = await AttachmentService.GetCountAsync("QMSDocument", Id.ToString());
                }
                catch
                {
                    attachmentCount = 0;
                }
            }
        }

        private Task<IEnumerable<EmployeeDropdownDto>> SearchEmployees(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return Task.FromResult<IEnumerable<EmployeeDropdownDto>>(employees.Take(100));
            }

            var query = searchText.Trim();
            var matches = employees.Where(e =>
                (!string.IsNullOrEmpty(e.EmpID) && e.EmpID.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(e.Name) && e.Name.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(e.FName) && e.FName.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(e.DeptName) && e.DeptName.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(e.Designation) && e.Designation.Contains(query, StringComparison.OrdinalIgnoreCase))
            );

            return Task.FromResult<IEnumerable<EmployeeDropdownDto>>(matches);
        }

        private void OnInitiatorSelected(EmployeeDropdownDto? emp)
        {
            selectedInitiator = emp;
            if (emp != null)
            {
                model.InitiaterEmpID = emp.EmpID;
                model.InitiatorName = emp.Name;
                model.InitiatorDepartment = emp.DeptName;
                model.InitiatorDesignation = emp.Designation;
            }
            else
            {
                model.InitiaterEmpID = null;
                model.InitiatorName = null;
                model.InitiatorDepartment = null;
                model.InitiatorDesignation = null;
            }
        }

        private void OnInternalScopeChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int val))
            {
                model.Internal = (val == 1);
            }
        }

        private void OnIssueDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out var dt))
            {
                model.IssueDT = dt;
            }
            else
            {
                model.IssueDT = null;
            }
        }

        private void AddEmployeeToDistribution()
        {
            if (selectedDistEmployee == null) return;

            bool alreadyExists = model.DistributionList.Any(d =>
                !d.IsMarkedForDeletion &&
                string.Equals(d.EmpID, selectedDistEmployee.EmpID, StringComparison.OrdinalIgnoreCase));

            if (alreadyExists)
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Already Added", $"Employee [{selectedDistEmployee.EmpID}] {selectedDistEmployee.Name} is already in the distribution list.");
                return;
            }

            model.DistributionList.Add(new QMSDocumentDistributionModel
            {
                EntryID = 0,
                QMS_RefID = model.EntryID,
                EmpID = selectedDistEmployee.EmpID,
                EmpName = selectedDistEmployee.Name,
                DeptName = selectedDistEmployee.DeptName,
                Designation = selectedDistEmployee.Designation,
                IsMarkedForDeletion = false
            });

            selectedDistEmployee = null;
        }

        private void RemoveEmployeeFromDistribution(QMSDocumentDistributionModel item)
        {
            if (item.EntryID > 0)
            {
                item.IsMarkedForDeletion = true;
                if (!deletedDistributionIds.Contains(item.EntryID))
                {
                    deletedDistributionIds.Add(item.EntryID);
                }
            }
            else
            {
                model.DistributionList.Remove(item);
            }
        }

        private void SetActiveTab(string tab)
        {
            activeTab = tab;
        }

        private async Task SaveAndClose()
        {
            await Save(closeAfterSave: true);
        }

        private async Task SaveAndNew()
        {
            await Save(closeAfterSave: false);
        }

        private async Task Save(bool closeAfterSave)
        {
            if (string.IsNullOrWhiteSpace(model.DocumentNumber))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Required Field", "Document Number is required.");
                return;
            }

            if (string.IsNullOrWhiteSpace(model.DocumentName))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Required Field", "Document Name / Title is required.");
                return;
            }

            if (string.IsNullOrWhiteSpace(model.DocumentType))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Required Field", "Document Type is required.");
                return;
            }

            if (!model.IssueDT.HasValue)
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Required Field", "Issue Date is required.");
                return;
            }

            try
            {
                isSaving = true;

                var authState = await AuthStateProvider.GetAuthenticationStateAsync();
                string currentUserName = authState.User?.Identity?.Name ?? "Admin";

                int savedId = await DocumentDataAccess.SaveDocumentAsync(model, deletedDistributionIds, currentUserName);
                model.EntryID = savedId;
                Id = savedId;

                NotificationService.Notify(NotificationSeverity.Success, "Saved", "Document saved successfully.");

                if (closeAfterSave)
                {
                    NavigateBack();
                }
                else
                {
                    NavigationManager.NavigateTo("/qms/documents/new?returnUrl=" + Uri.EscapeDataString(returnUrl ?? "/qms/documents"), forceLoad: true);
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Save Failed", $"Could not save document: {ex.Message}");
            }
            finally
            {
                isSaving = false;
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
                NavigationManager.NavigateTo("/qms/documents");
            }
        }
    }
}

