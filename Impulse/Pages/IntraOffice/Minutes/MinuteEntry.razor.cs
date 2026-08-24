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
        protected bool IsReadOnly => Id.HasValue && (MinuteModel.Status == "Approved" || MinuteModel.Status == "Rejected");

        protected string CurrentUserId { get; set; } = string.Empty;
        protected MinuteApproval MinuteModel { get; set; } = new();
        protected List<MinuteType> MinuteTypes { get; set; } = new();
        protected List<IntraUserProfile> UsersList { get; set; } = new();

        protected List<string> PointsList { get; set; } = new();
        protected string NewPointText { get; set; } = string.Empty;
        protected string DecisionRemarks { get; set; } = string.Empty;

        // Attachments dictionary mapped by slot index (0 to 7 -> I to VIII)
        protected Dictionary<int, IBrowserFile> PendingAttachments { get; set; } = new();

        protected override async Task OnInitializedAsync()
        {
            var authState = await AuthStateProvider.GetAuthenticationStateAsync();
            var user = authState.User;
            CurrentUserId = user.Identity?.Name ?? user.FindFirst(ClaimTypes.Name)?.Value ?? "Guest";

            await LoadDataAsync();
        }

        protected async Task LoadDataAsync()
        {
            IsLoading = true;
            try
            {
                MinuteTypes = await IntraOfficeService.GetMinuteTypesAsync();
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
                            Detail = "The requested minute memo does not exist.",
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
                        No = $"MIN-{DateTime.UtcNow:yyyyMMdd}-{new Random().Next(100, 999)}",
                        CreatedByUserId = CurrentUserId,
                        Status = "Pending",
                        Currency = "PKR"
                    };
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Load Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
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

        protected void HandlePointKeyUp(KeyboardEventArgs e)
        {
            if (e.Key == "Enter")
            {
                AddPoint();
            }
        }

        protected void CalculateAdvance()
        {
            if (MinuteModel.TotalAmount.HasValue && MinuteModel.AdvancePercentage.HasValue && MinuteModel.AdvancePercentage.Value > 0)
            {
                MinuteModel.AdvanceAmount = Math.Round((MinuteModel.TotalAmount.Value * MinuteModel.AdvancePercentage.Value) / 100m, 2);
            }
        }

        protected void HandleFileSelected(InputFileChangeEventArgs e, int slotIndex)
        {
            if (e.File != null)
            {
                PendingAttachments[slotIndex] = e.File;
            }
        }

        protected string GetRomanNumeral(int number)
        {
            string[] romans = { "I", "II", "III", "IV", "V", "VI", "VII", "VIII" };
            if (number >= 1 && number <= 8) return romans[number - 1];
            return number.ToString();
        }

        protected async Task SaveMinuteAsync()
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

                // Handle file uploads to wwwroot/uploads/minutes
                if (PendingAttachments.Any())
                {
                    var uploadFolder = System.IO.Path.Combine(WebHostEnvironment.WebRootPath ?? "wwwroot", "uploads", "minutes");
                    System.IO.Directory.CreateDirectory(uploadFolder);

                    foreach (var pair in PendingAttachments)
                    {
                        var file = pair.Value;
                        var uniqueName = $"{Guid.NewGuid():N}_{file.Name}";
                        var fullPath = System.IO.Path.Combine(uploadFolder, uniqueName);

                        await using (var fileStream = new System.IO.FileStream(fullPath, System.IO.FileMode.Create))
                        {
                            await file.OpenReadStream(maxAllowedSize: 20 * 1024 * 1024).CopyToAsync(fileStream);
                        }

                        MinuteModel.Attachments.Add(new MinuteAttachment
                        {
                            FileName = file.Name,
                            FilePath = $"/uploads/minutes/{uniqueName}",
                            FileSize = file.Size,
                            ContentType = file.ContentType
                        });
                    }
                }

                if (!Id.HasValue)
                {
                    var newId = await IntraOfficeService.CreateMinuteApprovalAsync(MinuteModel);
                    if (newId > 0)
                    {
                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Success,
                            Summary = "Created",
                            Detail = "Minute memo created and submitted successfully.",
                            Duration = 3000
                        });
                        NavigationManager.NavigateTo($"/intraoffice/minutes/{newId}");
                    }
                }
                else
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Saved",
                        Detail = "Minute memo updated successfully.",
                        Duration = 3000
                    });
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

        protected async Task ProcessDecisionAsync(string decision)
        {
            if (!Id.HasValue) return;

            try
            {
                var success = await IntraOfficeService.UpdateMinuteStatusAsync(
                    Id.Value,
                    decision,
                    CurrentUserId,
                    decision,
                    DecisionRemarks
                );

                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = decision == "Approved" ? NotificationSeverity.Success : NotificationSeverity.Warning,
                        Summary = decision,
                        Detail = $"Minute memo marked as {decision}.",
                        Duration = 3000
                    });

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
        }

        protected string GetStatusBadgeClass(string status)
        {
            return status switch
            {
                "Approved" => "bg-success",
                "Rejected" => "bg-danger",
                "Closed" => "bg-secondary",
                _ => "bg-warning text-dark"
            };
        }
    }
}
