namespace Impulse.Services
{
    public interface IReportManagerService
    {
        event Action<string, IDictionary<string, object>?> OnRequestReport;
        void RequestReport(string reportKey, IDictionary<string, object>? extraArgs = null);
    }
}
