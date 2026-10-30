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

namespace Impulse.Pages.QMS.CAPA
{
    public partial class CAPAEntry : ComponentBase
    {
        [Parameter] public int Id { get; set; }
        [SupplyParameterFromQuery(Name = "complaintId")] public int? ComplaintId { get; set; }
        [SupplyParameterFromQuery(Name = "ncId")] public int? NCId { get; set; }
        [SupplyParameterFromQuery(Name = "returnUrl")] public string? ReturnUrl { get; set; }

        [Inject] private ICustomerComplaintDataAccess ComplaintData { get; set; } = null!;
        [Inject] private IEntityAttachmentDataAccess AttachmentData { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNav { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private NotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IUserPermissionService PermissionService { get; set; } = null!;

        private bool IsEditMode => Id > 0;
        private string activeTab = "step1";
        private bool isLoading = true;
        private bool isSaving = false;
        private bool canCloseCapa = false;
        private bool showComplaintRef = true;
        private int attachmentCount = 0;

        private CAPAModel model = new();

        protected override async Task OnInitializedAsync()
        {
            isLoading = true;
            try
            {
                await PermissionService.InitializeAsync();
                canCloseCapa = PermissionService.IsAdministrator || PermissionService.HasOptionAccess("CloseCAPA");

                if (IsEditMode)
                {
                    var existing = await ComplaintData.GetCAPAByIdAsync(Id);
                    if (existing != null)
                    {
                        model = existing;
                        if (!ComplaintId.HasValue && model.ComplaintEntryID.HasValue)
                        {
                            ComplaintId = model.ComplaintEntryID;
                        }

                        var atts = await AttachmentData.GetAttachmentsAsync("CAPA", Id.ToString());
                        attachmentCount = atts.Count;
                    }
                    else
                    {
                        NotificationServiceManager.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Error,
                            Summary = "Not Found",
                            Detail = $"CAPA record #{Id} could not be found.",
                            Duration = 4000
                        });
                        Cancel();
                        return;
                    }
                }
                else
                {
                    var nextCapaNo = await ComplaintData.GetNextCAPANoAsync(DateTime.Today);
                    model = new CAPAModel
                    {
                        DT = DateTime.Today,
                        CAPANo = nextCapaNo,
                        Reason = 0,
                        Acceptance = 0,
                        CAReview = 0,
                        ResponseEvaluation = 0
                    };

                    if (ComplaintId.HasValue && ComplaintId.Value > 0)
                    {
                        var refData = await ComplaintData.GetCAPAByComplaintIdAsync(ComplaintId.Value);
                        if (refData != null)
                        {
                            // If this complaint already had a CAPA and opened in new mode inadvertently, load existing
                            if (refData.EntryID > 0)
                            {
                                Id = refData.EntryID;
                                model = refData;
                                var atts = await AttachmentData.GetAttachmentsAsync("CAPA", Id.ToString());
                                attachmentCount = atts.Count;
                                return;
                            }

                            // Copy complaint reference data
                            model.ComplaintEntryID = refData.ComplaintEntryID;
                            model.CustCode = refData.CustCode;
                            model.Country = refData.Country;
                            model.ComplaintNo = refData.ComplaintNo;
                            model.ComplaintDate = refData.ComplaintDate;
                            model.ComplaintBy = refData.ComplaintBy;
                            model.Designation = refData.Designation;
                            model.Penalty = refData.Penalty;
                            model.PenaltyQty = refData.PenaltyQty;
                            model.PenaltyAmount = refData.PenaltyAmount;
                            model.PenaltyCurrency = refData.PenaltyCurrency;
                            model.ComplaintReason = refData.ComplaintReason;
                            model.Description = refData.Description ?? refData.ComplaintReason;

                            // Copy NC reference data
                            model.NCEntryID = refData.NCEntryID;
                            model.NCNo = refData.NCNo;
                            model.NCType = refData.NCType;
                            model.NCDate = refData.NCDate;
                            model.NCSource = refData.NCSource;
                            model.NCDepartment = refData.NCDepartment;
                            model.NCDescription = refData.NCDescription;
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

        private void ToggleComplaintRef()
        {
            showComplaintRef = !showComplaintRef;
        }

        private async Task OnCAPADateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out var dt))
            {
                model.DT = dt;
                if (!IsEditMode)
                {
                    model.CAPANo = await ComplaintData.GetNextCAPANoAsync(dt);
                }
            }
        }

        private void HandleAttachmentCountChanged(int count)
        {
            attachmentCount = count;
        }

        private async Task SaveCAPA(bool saveAndNew)
        {
            if (string.IsNullOrWhiteSpace(model.CAPANo))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "CAPA # is required.",
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
                    Detail = "Non-Conformity Description is required.",
                    Duration = 3500
                });
                return;
            }

            isSaving = true;
            try
            {
                var userName = !string.IsNullOrWhiteSpace(PermissionService.CurrentUserName) ? PermissionService.CurrentUserName : "System";
                var savedId = await ComplaintData.SaveCAPAAsync(model, ComplaintId, userName);
                model.EntryID = savedId;
                Id = savedId;

                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "CAPA Saved",
                    Detail = $"CAPA {model.CAPANo} saved successfully.",
                    Duration = 3500
                });

                if (saveAndNew)
                {
                    var nextCapaNo = await ComplaintData.GetNextCAPANoAsync(DateTime.Today);
                    model = new CAPAModel
                    {
                        DT = DateTime.Today,
                        CAPANo = nextCapaNo,
                        Reason = 0,
                        Acceptance = 0,
                        CAReview = 0,
                        ResponseEvaluation = 0
                    };
                    Id = 0;
                    ComplaintId = null;
                    activeTab = "step1";
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

        private async Task PrintCAPA()
        {
            if (string.IsNullOrWhiteSpace(model.CAPANo))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "No CAPA",
                    Detail = "Please save CAPA before printing.",
                    Duration = 3500
                });
                return;
            }

            try
            {
                var request = new ReportRequest
                {
                    ReportName = ReportNames.QMS.CAPA,
                    SelectionFormula = $"{{CAPA.CAPANo}}='{model.CAPANo}'"
                };
                await ReportNav.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Print Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        private void Cancel()
        {
            NavigationManager.NavigateTo(ReturnUrl ?? "/qms/command-center");
        }
    }
}

