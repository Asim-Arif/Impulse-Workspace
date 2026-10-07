using System;
using System.Diagnostics;
using System.Net.Http;
using System.Threading.Tasks;
using Impulse.Services.IntraOffice;
using Microsoft.AspNetCore.Components;

namespace Impulse.Pages.IntraOffice.Admin
{
    public partial class HealthStatus : ComponentBase
    {
        [Inject] private IIntraOfficeService IntraOfficeService { get; set; } = null!;

        protected bool IsRefreshing { get; set; } = false;
        protected bool IsDbHealthy { get; set; } = true;
        protected long DbLatencyMs { get; set; } = 0;
        protected bool IsWhatsAppRunning { get; set; } = false;
        protected string WhatsAppStatusText { get; set; } = "Checking...";
        protected long ProcessMemoryMb { get; set; } = 0;
        protected string UptimeString { get; set; } = string.Empty;
        protected DateTime LastChecked { get; set; } = DateTime.UtcNow;

        protected override async Task OnInitializedAsync()
        {
            await RefreshHealthStatusAsync();
        }

        protected async Task RefreshHealthStatusAsync()
        {
            IsRefreshing = true;
            try
            {
                LastChecked = DateTime.UtcNow;

                // Check DB
                var sw = Stopwatch.StartNew();
                try
                {
                    var users = await IntraOfficeService.GetActiveUsersAsync();
                    sw.Stop();
                    DbLatencyMs = sw.ElapsedMilliseconds;
                    IsDbHealthy = true;
                }
                catch (Exception)
                {
                    IsDbHealthy = false;
                    DbLatencyMs = -1;
                }

                // Check Memory & Process
                var proc = Process.GetCurrentProcess();
                ProcessMemoryMb = proc.WorkingSet64 / (1024 * 1024);
                var uptime = DateTime.Now - proc.StartTime;
                UptimeString = $"{uptime.Days}d {uptime.Hours}h {uptime.Minutes}m";

                // Check WhatsApp Service
                try
                {
                    using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(2) };
                    var res = await http.GetAsync("http://127.0.0.1:3001/status");
                    if (res.IsSuccessStatusCode)
                    {
                        IsWhatsAppRunning = true;
                        WhatsAppStatusText = "Service Online";
                    }
                    else
                    {
                        IsWhatsAppRunning = false;
                        WhatsAppStatusText = "HTTP " + res.StatusCode;
                    }
                }
                catch
                {
                    IsWhatsAppRunning = false;
                    WhatsAppStatusText = "Offline (Local bridge stopped)";
                }
            }
            finally
            {
                IsRefreshing = false;
            }
        }
    }
}
