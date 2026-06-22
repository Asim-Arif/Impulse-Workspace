namespace Impulse.Services
{
    public interface IAuditService
    {
        string GetClientIpAddress();
        string GetCurrentUserName();
        //Task<string> GetCurrentUserName();
        // Add logging/saving methods here if needed:
        // Task LogTransaction(string action, string ip, string user);
    }
}
