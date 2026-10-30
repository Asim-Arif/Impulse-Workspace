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

namespace Impulse.Pages.QMS.NonCompliance
{
    public partial class NonComplianceEntry : ComponentBase
    {
        [Parameter] public int Id { get; set; }
        [SupplyParameterFromQuery(Name = "complaintId")] public int? ComplaintId { get; set; }
        [SupplyParameterFromQuery(Name = "returnUrl")] public string? ReturnUrl { get; set; }

        [Inject] private ICustomerComplaintDataAccess ComplaintData { get; set; } = null!;
        [Inject] private IEntityAttachmentDataAccess AttachmentData { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private NotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IUserPermissionService PermissionService { get; set; } = null!;

        private bool IsEditMode => Id > 0;
        private string activeTab = "details";
        private bool isLoading = true;
        private bool isSaving = false;
        private int attachmentCount = 0;

        private NonComplianceModel model = new();
        private List<DepartmentLookupDto> departments = new();

        protected override async Task OnInitializedAsync()
        {
            isLoading = true;
            try
            {
                await PermissionService.InitializeAsync();
                departments = await ComplaintData.GetDepartmentsAsync();

                if (IsEditMode)
                {
                    var existing = await ComplaintData.GetNCByIdAsync(Id);
                    if (existing != null)
                    {
                        model = existing;
                        if (!ComplaintId.HasValue && model.ComplaintEntryID.HasValue)
                        {
                            ComplaintId = model.ComplaintEntryID;
                        }

                        var atts = await AttachmentData.GetAttachmentsAsync("NonCompliance", Id.ToString());
                        attachmentCount = atts.Count;
                    }
                    else
                    {
                        NotificationServiceManager.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Error,
                            Summary = "Not Found",
                            Detail = $"Non-Compliance #{Id} could not be found.",
                            Duration = 4000
                        });
                        Cancel();
                        return;
                    }
                }
                else
                {
                    var nextNc = await ComplaintData.GetNextNCNoAsync(DateTime.Today);
                    model = new NonComplianceModel
                    {
                        DT = DateTime.Today,
                        NCNo = nextNc,
                        NCType = "Minor",
                        Source = "Customer Complaint"
                    };

                    if (ComplaintId.HasValue && ComplaintId.Value > 0)
                    {
                        var complaint = await ComplaintData.GetComplaintByIdAsync(ComplaintId.Value);
                        if (complaint != null)
                        {
                            model.ComplaintEntryID = complaint.EntryID;
                            model.CustCode = complaint.CustCode;
                            model.Country = complaint.Country;
                            model.ComplaintNo = complaint.ComplaintNo;
                            model.ComplaintDate = complaint.DT;
                            model.ComplaintBy = complaint.ComplaintBy;
                            model.Designation = complaint.Designation;
                            model.Penalty = complaint.Penalty;
                            model.PenaltyQty = complaint.PenaltyQty;
                            model.PenaltyAmount = complaint.PenaltyAmount;
                            model.PenaltyCurrency = complaint.PenaltyCurrency;
                            model.Description = complaint.Description;

                            model.Source = complaint.Type switch
                            {
                                1 => "Feedback/Observation",
                                2 => (complaint.InternalAudit == true ? "Internal Audit" : "External Audit"),
                                _ => "Customer Complaint"
                            };

                            if (!string.IsNullOrWhiteSpace(complaint.Unit))
                            {
                                model.Department = complaint.Unit;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Form",
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

        private async Task OnNCDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out var dt))
            {
                model.DT = dt;
                if (!IsEditMode)
                {
                    model.NCNo = await ComplaintData.GetNextNCNoAsync(dt);
                }
            }
        }

        private void HandleAttachmentCountChanged(int count)
        {
            attachmentCount = count;
        }

        private async Task SaveNC(bool saveAndNew)
        {
            if (string.IsNullOrWhiteSpace(model.NCNo))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "NC # is required.",
                    Duration = 3500
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(model.Department))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Department is required.",
                    Duration = 3500
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(model.Description))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Description of Non-Compliance is required.",
                    Duration = 3500
                });
                return;
            }

            isSaving = true;
            try
            {
                var userName = !string.IsNullOrWhiteSpace(PermissionService.CurrentUserName) ? PermissionService.CurrentUserName : "System";
                var savedId = await ComplaintData.SaveNCAsync(model, ComplaintId, userName);

                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "NC Saved",
                    Detail = $"Non-Compliance {model.NCNo} saved successfully.",
                    Duration = 3500
                });

                if (saveAndNew)
                {
                    var nextNc = await ComplaintData.GetNextNCNoAsync(DateTime.Today);
                    model = new NonComplianceModel
                    {
                        DT = DateTime.Today,
                        NCNo = nextNc,
                        NCType = "Minor",
                        Source = "Customer Complaint"
                    };
                    ComplaintId = null;
                    activeTab = "details";
                }
                else
                {
                    NavigationManager.NavigateTo(ReturnUrl ?? "/qms/command-center");
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
            NavigationManager.NavigateTo(ReturnUrl ?? "/qms/command-center");
        }
    }
}

