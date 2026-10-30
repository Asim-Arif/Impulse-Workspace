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
using Radzen;

namespace Impulse.Pages.QMS.DocumentsEx
{
    public partial class QMSDocumentExEntry : ComponentBase
    {
        [Parameter] public string? Category { get; set; }
        [Parameter] public int EntryId { get; set; }
        [SupplyParameterFromQuery] public string? ReturnUrl { get; set; }

        [Inject] private IQMSDocumentExDataAccess DocumentExData { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private NotificationService NotificationService { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private IUserPermissionService PermissionService { get; set; } = null!;

        private QMSDocumentExModel model = new();
        private List<QMSDocumentExDistributionItemModel> distributionList = new();
        private List<string> availableDocumentTypes = new();
        private List<EmployeeDropdownDto> allEmployees = new();

        private EmployeeDropdownDto? selectedInitiator;
        private EmployeeDropdownDto? selectedDistEmployee;

        private AttachmentModal? attachmentModalRef;
        private bool isLoading = true;
        private bool isSaving = false;

        private bool IsNew => EntryId <= 0;
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

        protected override async Task OnInitializedAsync()
        {
            await PermissionService.InitializeAsync();
            ResolveCategory();
            await LoadDataAsync();
        }

        private void ResolveCategory()
        {
            if (!string.IsNullOrWhiteSpace(Category))
            {
                var cat = Category.Trim().ToLowerInvariant();
                if (cat.Contains("manual")) MasterCategoryName = "Manuals";
                else if (cat.Contains("proc")) MasterCategoryName = "Procedures";
                else if (cat.Contains("work") || cat.Contains("instruction")) MasterCategoryName = "Work Instructions";
                else if (cat.Contains("tech") || cat.Contains("file")) MasterCategoryName = "Technical Files";
                else if (cat.Contains("lean")) MasterCategoryName = "Lean Management";
                else MasterCategoryName = "Standards";
            }
        }

        private async Task LoadDataAsync()
        {
            isLoading = true;
            try
            {
                allEmployees = await DocumentExData.GetEmployeesAsync();
                availableDocumentTypes = await DocumentExData.GetDocumentTypesByMasterAsync(MasterCategoryName);

                if (!IsNew)
                {
                    var loaded = await DocumentExData.GetDocumentByIdAsync(EntryId);
                    if (loaded != null)
                    {
                        model = loaded;
                        MasterCategoryName = loaded.Master_DocumentType;
                        distributionList = await DocumentExData.GetDistributionListAsync(EntryId);

                        if (!string.IsNullOrWhiteSpace(model.InitiaterEmpID))
                        {
                            selectedInitiator = allEmployees.FirstOrDefault(e => e.EmpID == model.InitiaterEmpID);
                        }
                    }
                    else
                    {
                        NotificationService.Notify(NotificationSeverity.Error, "Not Found", $"Document #{EntryId} not found.");
                        NavigateBack();
                        return;
                    }
                }
                else
                {
                    model = new QMSDocumentExModel
                    {
                        Master_DocumentType = MasterCategoryName,
                        IssueDT = DateTime.Today,
                        Revision = "00",
                        Internal = true
                    };
                    model.DocumentNumber = await DocumentExData.GetNextDocumentNumberAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error Loading", ex.Message);
            }
            finally
            {
                isLoading = false;
            }
        }

        private Task<IEnumerable<EmployeeDropdownDto>> SearchEmployees(string query)
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                return Task.FromResult<IEnumerable<EmployeeDropdownDto>>(allEmployees.Take(20));
            }

            var q = query.Trim();
            var matches = allEmployees.Where(e =>
                (!string.IsNullOrEmpty(e.EmpID) && e.EmpID.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(e.Name) && e.Name.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(e.FName) && e.FName.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(e.DeptName) && e.DeptName.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(e.Designation) && e.Designation.Contains(q, StringComparison.OrdinalIgnoreCase))
            ).Take(30);

            return Task.FromResult<IEnumerable<EmployeeDropdownDto>>(matches);
        }

        private void OnInitiatorSelected(EmployeeDropdownDto? emp)
        {
            selectedInitiator = emp;
            if (emp != null)
            {
                model.InitiaterEmpID = emp.EmpID;
                model.InitiatorName = emp.Name;
                model.InitiatorDept = emp.DeptName;
            }
            else
            {
                model.InitiaterEmpID = null;
                model.InitiatorName = null;
                model.InitiatorDept = null;
            }
        }

        private void AddEmployeeToDistribution()
        {
            if (selectedDistEmployee == null) return;

            bool alreadyExists = distributionList.Any(d =>
                !d.IsDeleted && string.Equals(d.EmpID, selectedDistEmployee.EmpID, StringComparison.OrdinalIgnoreCase));

            if (alreadyExists)
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Already Added", $"Employee [{selectedDistEmployee.EmpID}] {selectedDistEmployee.Name} is already in the distribution list.");
                return;
            }

            distributionList.Add(new QMSDocumentExDistributionItemModel
            {
                QMS_RefID = model.EntryID,
                EmpID = selectedDistEmployee.EmpID,
                Name = selectedDistEmployee.Name,
                DeptName = selectedDistEmployee.DeptName,
                Designation = selectedDistEmployee.Designation,
                IsNew = true,
                IsDeleted = false
            });

            selectedDistEmployee = null;
        }

        private void RemoveEmployeeFromDistribution(QMSDocumentExDistributionItemModel item)
        {
            if (item.EntryID > 0)
            {
                item.IsDeleted = true;
            }
            else
            {
                distributionList.Remove(item);
            }
        }

        private async Task GenerateDocNumber()
        {
            model.DocumentNumber = await DocumentExData.GetNextDocumentNumberAsync();
        }

        private async Task SaveDocumentAsync()
        {
            if (string.IsNullOrWhiteSpace(model.DocumentNumber))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation Warning", "Document Number is required.");
                return;
            }

            if (string.IsNullOrWhiteSpace(model.DocumentName))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation Warning", "Document Title/Name is required.");
                return;
            }

            model.Master_DocumentType = MasterCategoryName;

            isSaving = true;
            try
            {
                int savedId = await DocumentExData.SaveDocumentAsync(model, distributionList);
                model.EntryID = savedId;

                NotificationService.Notify(NotificationSeverity.Success, "Saved Successfully", $"Document [{model.DocumentNumber}] saved successfully.");
                NavigateBack();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Save Failed", ex.Message);
            }
            finally
            {
                isSaving = false;
            }
        }

        private void NavigateBack()
        {
            if (!string.IsNullOrWhiteSpace(ReturnUrl))
            {
                NavigationManager.NavigateTo(ReturnUrl);
            }
            else
            {
                NavigationManager.NavigateTo($"/qms/{CategorySlug}");
            }
        }

        private void OpenAttachments()
        {
            if (model.EntryID > 0)
            {
                attachmentModalRef?.Open(model.EntryID.ToString(), "QMSDocumentEx", model.DocumentName);
            }
        }

        private async Task PrintDistributionList()
        {
            if (model.EntryID <= 0) return;

            try
            {
                string selectionFormula = $"{{QMS_Documents_Ex.EntryID}} = {model.EntryID}";
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
    }
}

