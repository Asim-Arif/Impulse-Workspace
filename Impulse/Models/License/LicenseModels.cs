using System;
using System.Collections.Generic;

namespace Impulse.Models.License
{
    [Flags]
    public enum AppModule
    {
        None = 0,
        IntraOffice = 1,
        Payroll = 2,
        Production = 4,
        Stock = 8,
        Accounts = 16,
        Export = 32,
        All = IntraOffice | Payroll | Production | Stock | Accounts | Export
    }

    public enum LicenseStatus
    {
        Valid,
        Missing,
        InvalidSignature,
        HardwareMismatch,
        Expired,
        Corrupted
    }

    public class LicensePayload
    {
        public string LicenseId { get; set; } = Guid.NewGuid().ToString("N");
        public string ClientName { get; set; } = string.Empty;
        public string HardwareId { get; set; } = string.Empty; // HWID-XXXX-XXXX-XXXX-XXXX or "*"
        public List<string> AllowedModules { get; set; } = new();
        public DateTime IssuedAt { get; set; } = DateTime.UtcNow;
        public DateTime? ExpiresAt { get; set; } // Null = Lifetime
        public int MaxUsers { get; set; } = 0;   // 0 = Unlimited
    }

    public class LicenseInfo
    {
        public LicenseStatus Status { get; set; } = LicenseStatus.Missing;
        public bool IsLicensed => Status == LicenseStatus.Valid;
        public string ClientName { get; set; } = string.Empty;
        public string HardwareId { get; set; } = string.Empty;
        public HashSet<string> AllowedModules { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        public DateTime? ExpiresAt { get; set; }
        public DateTime? IssuedAt { get; set; }
        public string StatusMessage { get; set; } = "No license found.";

        public bool IsModuleAllowed(AppModule module)
        {
            if (!IsLicensed) return false;
            if (AllowedModules.Contains("All")) return true;
            return AllowedModules.Contains(module.ToString());
        }

        public bool IsModuleAllowed(string moduleName)
        {
            if (!IsLicensed) return false;
            if (AllowedModules.Contains("All")) return true;
            if (string.IsNullOrWhiteSpace(moduleName)) return true;

            var lower = moduleName.Trim().ToLowerInvariant();

            // Core system areas are part of the base application and accessible in all licensed tiers
            if (lower is "setup" or "setups" or "dashboard" or "dashboards" or "company" or "favourites" or "help")
            {
                return true;
            }

            var normalized = lower switch
            {
                "financial" or "accounts" => "Accounts",
                "payroll" or "hr" => "Payroll",
                "stock" or "inventory" => "Stock",
                "production" => "Production",
                "export" => "Export",
                "office" or "intraoffice" => "IntraOffice",
                _ => moduleName.Trim()
            };

            return AllowedModules.Contains(normalized) || AllowedModules.Contains(moduleName);
        }
    }

    public class LicenseDbRecord
    {
        public int EntryID { get; set; }
        public string? LicenseKey { get; set; }
        public string? MachineName { get; set; }
    }
}
