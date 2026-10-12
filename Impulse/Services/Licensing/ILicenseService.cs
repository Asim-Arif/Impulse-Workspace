using System.Threading.Tasks;
using Impulse.Models.License;

namespace Impulse.Services.Licensing
{
    public interface ILicenseService
    {
        LicenseInfo CurrentLicense { get; }
        bool IsLicensed { get; }
        bool IsModuleAllowed(AppModule module);
        bool IsModuleAllowed(string moduleName);
        string GetCurrentHardwareId();
        Task ReloadLicenseAsync();
        Task<(bool Success, string Message)> ActivateLicenseAsync(string newLicenseKey);
    }
}
