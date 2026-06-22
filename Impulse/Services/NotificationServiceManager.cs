using Microsoft.JSInterop;
using Radzen;

namespace CIP.Services
{
    public class NotificationServiceManager : INotificationService
    {
        private readonly NotificationService _notificationService;
        private readonly DialogService _dialogService;
        private readonly double dDuration = 5000;
        private readonly IJSRuntime _js;
        public NotificationServiceManager(NotificationService notificationService, IJSRuntime js, DialogService dialogService)
        {
            _notificationService = notificationService;
            _js = js;
            _dialogService = dialogService;
        }
        private string GetCustomStyle(string bgColor, string textColor, string borderColor)
        {
            // Replace with your actual MsgStyle logic or return the custom CSS string
            return $"background-color: {bgColor}; color: {textColor}; border-color: {borderColor};";
        }

        public void ShowWarning(string summary, string detail)
        {
            _notificationService.Notify(new NotificationMessage
            {
                Severity = NotificationSeverity.Warning,
                Summary = $"{summary}",
                Detail = detail,
                Duration = dDuration,
                Style = GetCustomStyle("#fff3cd", "#856404", "#ffc107")
            });
        }

        public void ShowSuccess(string summary, string detail)
        {
            _notificationService.Notify(new NotificationMessage
            {
                Severity = NotificationSeverity.Success,
                Summary = $"{summary}",
                Detail = detail,
                Duration = dDuration,
                Style = GetCustomStyle("#fff3cd", "#856404", "#ffc107")
            });
        }

        public void ShowError(string summary, string detail)
        {
            _notificationService.Notify(new NotificationMessage
            {
                Severity = NotificationSeverity.Error,
                Summary = summary,
                Detail = detail,
                Duration = dDuration,
                Style = GetCustomStyle("#f8d7da", "#721c24", "#dc3545")
            });
        }
        public void ShowInformation(string summary, string detail)
        {
            _notificationService.Notify(new NotificationMessage
            {
                Severity = NotificationSeverity.Info,
                Summary = summary,
                Detail = detail,
                Duration = dDuration,
                Style = GetCustomStyle("#f8d7da", "#721c24", "#dc3545")
            });
        }

        public async Task<bool> ShowQuestionNotification(string summary, string detail)
        {
            return await _js.InvokeAsync<bool>("confirm", $"{summary}\n\n{detail}");

        }
    }
}
