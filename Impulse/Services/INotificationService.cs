namespace Impulse.Services
{
    public interface INotificationService
    {
        void ShowSuccess(string summary,string detail);
        void ShowWarning(string summary, string detail);
        void ShowError(string summary, string detail);
        void ShowInformation(string summary, string detail);
        Task<bool> ShowQuestionNotification(string summary, string detail);
    }
}
