using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Extensions.Logging;
using Microsoft.Win32;

namespace Impulse.Services.Licensing
{
    public class HardwareFingerprintService : IHardwareFingerprintService
    {
        private readonly ILogger<HardwareFingerprintService> _logger;
        private string? _cachedHwid;

        public HardwareFingerprintService(ILogger<HardwareFingerprintService> logger)
        {
            _logger = logger;
        }

        public string GetHardwareId()
        {
            if (!string.IsNullOrEmpty(_cachedHwid))
            {
                return _cachedHwid;
            }

            try
            {
                var sb = new StringBuilder();

                // 1. Windows Machine GUID
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    try
                    {
                        var machineGuid = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Cryptography", "MachineGuid", null)?.ToString();
                        if (!string.IsNullOrWhiteSpace(machineGuid))
                        {
                            sb.Append("MG:").Append(machineGuid.Trim());
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Could not read MachineGuid from registry");
                    }
                }

                // 2. Machine Hostname
                sb.Append("|MN:").Append(Environment.MachineName.Trim().ToUpperInvariant());

                // 3. System drive volume / Windows Directory info
                try
                {
                    var winDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
                    if (!string.IsNullOrEmpty(winDir))
                    {
                        var driveInfo = new DriveInfo(Path.GetPathRoot(winDir) ?? "C:\\");
                        sb.Append("|DV:").Append(driveInfo.Name).Append(":").Append(driveInfo.DriveFormat);
                    }
                }
                catch { }

                // 4. Processor count + OS architecture
                sb.Append("|PR:").Append(Environment.ProcessorCount)
                  .Append("|OS:").Append(RuntimeInformation.OSArchitecture);

                var rawData = sb.ToString();

                using var sha = SHA256.Create();
                var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(rawData));

                // Format as clean 16-character alphanumeric code: HWID-XXXX-XXXX-XXXX-XXXX
                var hex = Convert.ToHexString(hash);
                _cachedHwid = $"HWID-{hex.Substring(0, 4)}-{hex.Substring(4, 4)}-{hex.Substring(8, 4)}-{hex.Substring(12, 4)}";
                return _cachedHwid;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating hardware fingerprint. Using fallback.");
                return $"HWID-{Environment.MachineName.ToUpperInvariant().PadRight(16, '0').Substring(0, 16)}";
            }
        }

        public bool MatchesHardwareId(string targetHardwareId)
        {
            if (string.IsNullOrWhiteSpace(targetHardwareId)) return false;
            var trimmed = targetHardwareId.Trim();
            if (trimmed == "*") return true; // Wildcard / cloud license

            var current = GetHardwareId();
            return string.Equals(trimmed, current, StringComparison.OrdinalIgnoreCase);
        }
    }
}
