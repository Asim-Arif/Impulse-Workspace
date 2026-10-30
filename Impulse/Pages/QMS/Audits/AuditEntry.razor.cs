using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Common;
using DataAccessLibrary.Interface.QMS;
using DataAccessLibrary.Models.QMS;
using Impulse.Services.Setup;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.QMS.Audits
{
    public partial class AuditEntry : ComponentBase
    {
        [Parameter] public int Id { get; set; }
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

        private CustomerComplaintModel model = new();

        protected override async Task OnInitializedAsync()
        {
            isLoading = true;
            try
            {
                await PermissionService.InitializeAsync();

                if (IsEditMode)
                {
                    var existing = await ComplaintData.GetComplaintByIdAsync(Id);
                    if (existing != null)
                    {
                        model = existing;
                        var atts = await AttachmentData.GetAttachmentsAsync("CustomerComplaint", Id.ToString());
                        attachmentCount = atts.Count;
                    }
                    else
                    {
                        NotificationServiceManager.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Error,
                            Summary = "Not Found",
                            Detail = $"Audit #{Id} could not be found.",
                            Duration = 4000
                        });
                        Cancel();
                        return;
                    }
                }
                else
                {
                    model = new CustomerComplaintModel
                    {
                        Type = 2,
                        InternalAudit = true,
                        DT = DateTime.Today,
                        ComplaintNo = await ComplaintData.GetNextAuditNoAsync(DateTime.Today)
                    };
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

        private void SetAuditType(bool isInternal)
        {
            model.InternalAudit = isInternal;
        }

        private async Task OnAuditDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out var dt))
            {
                model.DT = dt;
                if (!IsEditMode)
                {
                    model.ComplaintNo = await ComplaintData.GetNextAuditNoAsync(dt);
                }
            }
        }

        private void OnNextAuditDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out var dt))
            {
                model.DTNextAudit = dt;
            }
            else
            {
                model.DTNextAudit = null;
            }
        }

        private void HandleAttachmentCountChanged(int count)
        {
            attachmentCount = count;
        }

        private async Task SaveAudit(bool saveAndNew)
        {
            if (string.IsNullOrWhiteSpace(model.ComplaintNo))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Audit # is required.",
                    Duration = 3500
                });
                return;
            }

            isSaving = true;
            try
            {
                model.Type = 2; // 2: Internal / External Audit
                var userName = !string.IsNullOrWhiteSpace(PermissionService.CurrentUserName) ? PermissionService.CurrentUserName : "System";
                var machineName = Environment.MachineName;

                var savedId = await ComplaintData.SaveComplaintAsync(model, model.Details, userName, machineName);

                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Audit Saved",
                    Detail = $"Audit {model.ComplaintNo} saved successfully.",
                    Duration = 3500
                });

                if (saveAndNew)
                {
                    model = new CustomerComplaintModel
                    {
                        Type = 2,
                        InternalAudit = true,
                        DT = DateTime.Today,
                        ComplaintNo = await ComplaintData.GetNextAuditNoAsync(DateTime.Today)
                    };
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

