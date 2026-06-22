namespace CIP.Services
{
    public class ReportManagerService : IReportManagerService
    {
        // This event tells MainLayout to show a specific report modal
        public event Action<string, IDictionary<string, object>?> OnRequestReport;
        public void RequestReport(string reportKey, IDictionary<string, object>? extraArgs = null)
        {
            // Invoke the event with the key and the dictionary (which might be null)
            OnRequestReport?.Invoke(reportKey, extraArgs);
        }
    }
}
