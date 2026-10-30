using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Common;
using DataAccessLibrary.Interface.QMS;
using DataAccessLibrary.Models.QMS;
using Impulse.Constants;
using Impulse.Services;
using Impulse.Services.Setup;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.QMS.CreditNotes
{
    public partial class CustomerCreditNoteEntry : ComponentBase
    {
        [Parameter] public int ComplaintId { get; set; }
        [SupplyParameterFromQuery(Name = "complaintId")] public int? QueryComplaintId { get; set; }
        [SupplyParameterFromQuery(Name = "returnUrl")] public string? ReturnUrl { get; set; }

        [Inject] private ICustomerComplaintDataAccess ComplaintData { get; set; } = null!;
        [Inject] private IEntityAttachmentDataAccess AttachmentData { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNav { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private NotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IUserPermissionService PermissionService { get; set; } = null!;

        private int effectiveComplaintId => ComplaintId > 0 ? ComplaintId : (QueryComplaintId ?? 0);

        private string activeTab = "details";
        private bool isLoading = true;
        private bool isSaving = false;
        private int attachmentCount = 0;

        private string creditNoteNo = string.Empty;
        private DateTime creditNoteDate = DateTime.Today;
        private CustomerComplaintModel? complaint;

        private double TotalAmount => complaint?.Details
            .Where(d => !d.IsDeleted)
            .Sum(d => (double)(d.Qty ?? 0) * (double)(d.Price ?? 0)) ?? 0;

        protected override async Task OnInitializedAsync()
        {
            isLoading = true;
            try
            {
                await PermissionService.InitializeAsync();

                if (effectiveComplaintId > 0)
                {
                    complaint = await ComplaintData.GetComplaintByIdAsync(effectiveComplaintId);
                    if (complaint != null)
                    {
                        if (!string.IsNullOrWhiteSpace(complaint.CreditNoteNo))
                        {
                            creditNoteNo = complaint.CreditNoteNo;
                        }
                        else
                        {
                            creditNoteNo = await ComplaintData.GetNextCreditNoteNoAsync(DateTime.Today);
                        }

                        var atts = await AttachmentData.GetAttachmentsAsync("CustomerComplaint", effectiveComplaintId.ToString());
                        attachmentCount = atts.Count;
                    }
                    else
                    {
                        NotificationServiceManager.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Error,
                            Summary = "Not Found",
                            Detail = $"Complaint #{effectiveComplaintId} could not be found.",
                            Duration = 4000
                        });
                        Cancel();
                        return;
                    }
                }
                else
                {
                    NotificationServiceManager.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Warning,
                        Summary = "Missing Complaint",
                        Detail = "Please select a customer complaint to generate a credit note.",
                        Duration = 4000
                    });
                    Cancel();
                    return;
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

        private async Task OnCreditNoteDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out var dt))
            {
                creditNoteDate = dt;
                if (complaint != null && string.IsNullOrWhiteSpace(complaint.CreditNoteNo))
                {
                    creditNoteNo = await ComplaintData.GetNextCreditNoteNoAsync(dt);
                }
            }
        }

        private void OnPriceChanged(CustomerComplaintDetailModel detail, ChangeEventArgs e)
        {
            if (float.TryParse(e.Value?.ToString(), out var price))
            {
                detail.Price = price;
            }
            else
            {
                detail.Price = 0;
            }
        }

        private void HandleAttachmentCountChanged(int count)
        {
            attachmentCount = count;
        }

        private async Task SaveCreditNote()
        {
            if (string.IsNullOrWhiteSpace(creditNoteNo))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Credit Note # is required.",
                    Duration = 3500
                });
                return;
            }

            if (complaint == null) return;

            isSaving = true;
            try
            {
                var userName = !string.IsNullOrWhiteSpace(PermissionService.CurrentUserName) ? PermissionService.CurrentUserName : "System";
                await ComplaintData.SaveCreditNoteAsync(effectiveComplaintId, creditNoteNo.Trim(), complaint.Details, userName);

                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Credit Note Saved",
                    Detail = $"Credit Note {creditNoteNo} saved successfully. Total: {TotalAmount:N2}",
                    Duration = 3500
                });

                NavigationManager.NavigateTo(ReturnUrl ?? "/qms/command-center");
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

        private async Task PrintCreditNote()
        {
            if (effectiveComplaintId <= 0) return;

            var request = new ReportRequest
            {
                ReportName = ReportNames.QMS.CreditNote,
                SelectionFormula = $"{{CustomerComplaints.EntryID}} = {effectiveComplaintId} AND {{CustomerComplaints_Detail.Price}} <> 0"
            };
            await ReportNav.PrintReportAsync(request);
        }

        private void Cancel()
        {
            NavigationManager.NavigateTo(ReturnUrl ?? "/qms/command-center");
        }
    }
}

