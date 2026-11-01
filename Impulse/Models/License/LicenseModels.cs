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
        
        // Add-on Modules (licensed explicitly, excluded from standard "All")
        QMS = 64,
        Budgeting = 128,
        FixedAssets = 256,
        GateManagement = 512,

        // Core ERP represents the standard "All" bundle
        CoreAll = IntraOffice | Payroll | Production | Stock | Accounts | Export,
        All = CoreAll
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

        public static bool IsAddonModule(string moduleName)
        {
            var lower = moduleName.Trim().ToLowerInvariant().Replace(" ", "").Replace("-", "");
            return lower switch
            {
                "qms" or "quality" => true,
                "budgeting" or "budget" => true,
                "fixedassets" => true,
                "gate" or "gatemanagement" => true,
                _ => false
            };
        }

        public static bool IsAddonModule(AppModule module)
        {
            return module == AppModule.QMS || module == AppModule.Budgeting || module == AppModule.FixedAssets || module == AppModule.GateManagement;
        }

        public bool IsModuleAllowed(AppModule module)
        {
            if (!IsLicensed) return false;
            if (AllowedModules.Contains("EnterpriseAll") || AllowedModules.Contains("Everything")) return true;
            if (AllowedModules.Contains(module.ToString())) return true;
            if (AllowedModules.Contains("All"))
            {
                return !IsAddonModule(module);
            }
            return false;
        }

        public bool IsModuleAllowed(string moduleName)
        {
            if (!IsLicensed) return false;
            if (string.IsNullOrWhiteSpace(moduleName)) return true;

            var lower = moduleName.Trim().ToLowerInvariant().Replace(" ", "").Replace("-", "");

            // Core system areas are part of the base application and accessible in all licensed tiers
            if (lower is "setup" or "setups" or "dashboard" or "dashboards" or "company" or "favourites" or "help")
            {
                return true;
            }

            var normalized = lower switch
            {
                "financial" or "accounts" => "Accounts",
                "budgeting" or "budget" => "Budgeting",
                "fixedassets" => "FixedAssets",
                "gate" or "gatemanagement" => "GateManagement",
                "payroll" or "hr" => "Payroll",
                "stock" or "inventory" => "Stock",
                "production" => "Production",
                "export" => "Export",
                "office" or "intraoffice" => "IntraOffice",
                "qms" or "quality" => "QMS",
                _ => moduleName.Trim()
            };

            // 1. Explicitly licensed module (e.g., license contains "QMS")
            if (AllowedModules.Contains(normalized) || AllowedModules.Contains(moduleName.Trim()))
            {
                return true;
            }

            // 2. Full enterprise / developer wildcard
            if (AllowedModules.Contains("EnterpriseAll") || AllowedModules.Contains("Everything"))
            {
                return true;
            }

            // 3. Standard "All" covers only Core ERP modules, NOT Add-ons (like QMS)
            if (AllowedModules.Contains("All"))
            {
                return !IsAddonModule(normalized);
            }

            return false;
        }
    }

    public class LicenseDbRecord
    {
        public int EntryID { get; set; }
        public string? LicenseKey { get; set; }
        public string? MachineName { get; set; }
    }
}
