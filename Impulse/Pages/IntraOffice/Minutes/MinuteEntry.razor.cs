using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using DataAccessLibrary.Models.IntraOffice;
using Impulse.Services.IntraOffice;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Hosting;
using Radzen;

namespace Impulse.Pages.IntraOffice.Minutes
{
    public partial class MinuteEntry : ComponentBase
    {
        [Parameter] public int? Id { get; set; }

        [Inject] private IIntraOfficeService IntraOfficeService { get; set; } = null!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private NotificationService NotificationService { get; set; } = null!;
        [Inject] private IWebHostEnvironment WebHostEnvironment { get; set; } = null!;

        protected bool IsLoading { get; set; } = true;
        protected bool IsSaving { get; set; } = false;
        protected bool IsEditMode { get; set; } = false;

        protected string CurrentUserId { get; set; } = string.Empty;
        protected MinuteApproval MinuteModel { get; set; } = new();
        protected List<IntraUserProfile> UsersList { get; set; } = new();

        protected List<string> PointsList { get; set; } = new();
        protected string NewPointText { get; set; } = string.Empty;

        // Attachments
        protected Dictionary<int, IBrowserFile> AttachmentFiles { get; set; } = new();
        protected string[] AttachmentNames { get; set; } = new string[8];

        // Workflow Action fields
        protected string NextForwardToUserId { get; set; } = string.Empty;
        protected string DecisionRemarks { get; set; } = string.Empty;

        protected override async Task OnInitializedAsync()
        {
            var authState = await AuthStateProvider.GetAuthenticationStateAsync();
            var user = authState.User;
            CurrentUserId = user.Identity?.Name ?? user.FindFirst(ClaimTypes.Name)?.Value ?? "Guest";

            for (int i = 0; i < 8; i++)
            {
                AttachmentNames[i] = string.Empty;
            }

            await LoadDataAsync();
        }

        protected async Task LoadDataAsync()
        {
            IsLoading = true;
            try
            {
                UsersList = await IntraOfficeService.GetActiveUsersAsync();

                if (Id.HasValue && Id.Value > 0)
                {
                    var existing = await IntraOfficeService.GetMinuteByIdAsync(Id.Value);
                    if (existing != null)
                    {
                        MinuteModel = existing;
                        if (!string.IsNullOrWhiteSpace(existing.Points))
                        {
                            PointsList = existing.Points.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries).ToList();
                        }
                    }
                    else
                    {
                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Error,
                            Summary = "Not Found",
                            Detail = "The requested minute memo was not found.",
                            Duration = 4000
                        });
                        NavigationManager.NavigateTo("/intraoffice/minutes");
                    }
                }
                else
                {
                    MinuteModel = new MinuteApproval
                    {
                        Date = DateTime.Today,
                        No = $"{DateTime.Today.Month}/{new Random().Next(100, 9999):D4}",
                        CreatedByUserId = CurrentUserId,
                        Status = "Pending",
                        Type = "General"
                    };
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        protected void HandlePointKeyPress(KeyboardEventArgs e)
        {
            if (e.Key == "Enter")
            {
                AddPoint();
            }
        }

        protected void AddPoint()
        {
            if (!string.IsNullOrWhiteSpace(NewPointText))
            {
                PointsList.Add(NewPointText.Trim());
                NewPointText = string.Empty;
            }
        }

        protected void HandleFileSelected(InputFileChangeEventArgs e, int index)
        {
            if (e.File != null)
            {
                AttachmentFiles[index] = e.File;
                if (string.IsNullOrWhiteSpace(AttachmentNames[index]))
                {
                    AttachmentNames[index] = e.File.Name;
                }
            }
        }

        protected void RemoveAttachmentSlot(int index)
        {
            AttachmentFiles.Remove(index);
            AttachmentNames[index] = string.Empty;
        }

        protected string GetRomanNumeral(int number)
        {
            string[] romans = { "I", "II", "III", "IV", "V", "VI", "VII", "VIII" };
            if (number >= 1 && number <= 8) return romans[number - 1];
            return number.ToString();
        }

        protected async Task SaveMinuteWithAction(string action)
        {
            if (string.IsNullOrWhiteSpace(MinuteModel.No))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Please enter a minute number.",
                    Duration = 3000
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(MinuteModel.Subject))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Please enter a subject.",
                    Duration = 3000
                });
                return;
            }

            IsSaving = true;
            try
            {
                MinuteModel.Points = string.Join("\n", PointsList);
                MinuteModel.CreatedByUserId = CurrentUserId;

                // Handle file uploads to wwwroot/uploads/minutes
                if (AttachmentFiles.Any())
                {
                    var uploadFolder = System.IO.Path.Combine(WebHostEnvironment.WebRootPath ?? "wwwroot", "uploads", "minutes");
                    System.IO.Directory.CreateDirectory(uploadFolder);

                    foreach (var pair in AttachmentFiles)
                    {
                        var slotIndex = pair.Key;
                        var file = pair.Value;
                        var customTitle = !string.IsNullOrWhiteSpace(AttachmentNames[slotIndex]) ? AttachmentNames[slotIndex].Trim() : file.Name;
                        var uniqueName = $"{Guid.NewGuid():N}_{file.Name}";
                        var fullPath = System.IO.Path.Combine(uploadFolder, uniqueName);

                        await using (var fileStream = new System.IO.FileStream(fullPath, System.IO.FileMode.Create))
                        {
                            await file.OpenReadStream(maxAllowedSize: 25 * 1024 * 1024).CopyToAsync(fileStream);
                        }

                        MinuteModel.Attachments.Add(new MinuteAttachment
                        {
                            FileName = customTitle,
                            FilePath = $"/uploads/minutes/{uniqueName}",
                            FileSize = file.Size,
                            ContentType = file.ContentType
                        });
                    }
                }

                var newId = await IntraOfficeService.CreateMinuteApprovalAsync(MinuteModel);
                if (newId > 0)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Success",
                        Detail = "Minute created successfully.",
                        Duration = 3000
                    });

                    if (action == "new")
                    {
                        MinuteModel = new MinuteApproval
                        {
                            Date = DateTime.Today,
                            No = $"{DateTime.Today.Month}/{new Random().Next(100, 9999):D4}",
                            CreatedByUserId = CurrentUserId,
                            Status = "Pending",
                            Type = "General"
                        };
                        PointsList.Clear();
                        AttachmentFiles.Clear();
                        for (int i = 0; i < 8; i++) AttachmentNames[i] = string.Empty;
                    }
                    else if (action == "close")
                    {
                        NavigationManager.NavigateTo("/intraoffice/minutes");
                    }
                    else // stay
                    {
                        Id = newId;
                        await LoadDataAsync();
                    }
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Save Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsSaving = false;
            }
        }

        protected async Task ProcessAction(string action)
        {
            if (!Id.HasValue) return;

            IsSaving = true;
            try
            {
                var targetStatus = action switch
                {
                    "Approve" => "Approved",
                    "Reject" => "Rejected",
                    "Case Close" => "Closed",
                    "Refer" => "Referred",
                    _ => "Pending"
                };

                var success = await IntraOfficeService.UpdateMinuteStatusAsync(
                    Id.Value,
                    targetStatus,
                    CurrentUserId,
                    action,
                    DecisionRemarks
                );

                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = targetStatus == "Approved" ? NotificationSeverity.Success : NotificationSeverity.Info,
                        Summary = action,
                        Detail = $"Minute has been {targetStatus}.",
                        Duration = 3000
                    });

                    DecisionRemarks = string.Empty;
                    await LoadDataAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsSaving = false;
            }
        }

        protected void Cancel()
        {
            NavigationManager.NavigateTo("/intraoffice/minutes");
        }

        protected string GetStatusBadgeClass(string status)
        {
            return status switch
            {
                "Approved" => "bg-success",
                "Rejected" => "bg-danger",
                "Closed" => "bg-secondary",
                "Referred" => "bg-warning text-dark",
                _ => "bg-primary"
            };
        }
    }
}
