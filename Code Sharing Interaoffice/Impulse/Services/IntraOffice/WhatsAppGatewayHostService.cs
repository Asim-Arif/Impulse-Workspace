using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.IntraOffice
{
    public class WhatsAppGatewayHostService : BackgroundService
    {
        private readonly ILogger<WhatsAppGatewayHostService> _logger;
        private Process? _nodeProcess;

        public WhatsAppGatewayHostService(ILogger<WhatsAppGatewayHostService> logger)
        {
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _logger.LogInformation("WhatsApp Gateway Host Manager background service starting...");

            // Initial brief delay to let Kestrel finish startup
            await Task.Delay(TimeSpan.FromSeconds(3), stoppingToken);

            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    bool isAlive = await CheckIfServiceAliveAsync();
                    if (!isAlive)
                    {
                        _logger.LogWarning("WhatsApp Gateway (127.0.0.1:3001) is offline. Attempting to start/restart microservice...");
                        EnsureNodeProcessRunning();
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error in WhatsAppGatewayHostService supervision loop.");
                }

                // Check health every 20 seconds
                await Task.Delay(TimeSpan.FromSeconds(20), stoppingToken);
            }

            CleanupProcess();
        }

        private async Task<bool> CheckIfServiceAliveAsync()
        {
            try
            {
                using var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(3);
                var response = await client.GetAsync("http://127.0.0.1:3001/status");
                return response.IsSuccessStatusCode;
            }
            catch
            {
                return false;
            }
        }

        private void EnsureNodeProcessRunning()
        {
            try
            {
                if (_nodeProcess != null && !_nodeProcess.HasExited)
                {
                    _logger.LogInformation("Node process is already tracked (PID {Pid}), awaiting ready state...", _nodeProcess.Id);
                    return;
                }

                // Find whatsapp-service directory across standard and relative deployment locations
                string[] candidateDirs = new[]
                {
                    Path.Combine(AppContext.BaseDirectory, "whatsapp-service"),
                    Path.Combine(Directory.GetCurrentDirectory(), "whatsapp-service"),
                    Path.Combine(Directory.GetCurrentDirectory(), "Impulse", "whatsapp-service"),
                    Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "whatsapp-service"),
                    Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "InteraOffice Latest", "Source", "whatsapp-service"),
                    Path.Combine(Directory.GetCurrentDirectory(), "InteraOffice Latest", "Source", "whatsapp-service"),
                    Path.Combine(Directory.GetCurrentDirectory(), "..", "InteraOffice Latest", "Source", "whatsapp-service"),
                    @"C:\Visual Studio .Net\Impulse Workspace\Impulse Solution\Impulse\whatsapp-service",
                    @"C:\Visual Studio .Net\Impulse Workspace\InteraOffice Latest\Source\whatsapp-service",
                    @"D:\IntraCom-CRM\whatsapp-service",
                    @"D:\Intra Office\whatsapp-service"
                };

                string? serviceDir = candidateDirs.FirstOrDefault(d => Directory.Exists(d) && File.Exists(Path.Combine(d, "server.js")));
                if (serviceDir == null)
                {
                    _logger.LogWarning("Could not find whatsapp-service directory containing server.js.");
                    return;
                }

                // Auto-restore node_modules if missing
                if (!Directory.Exists(Path.Combine(serviceDir, "node_modules")))
                {
                    _logger.LogInformation("node_modules missing in {Dir}. Running npm install...", serviceDir);
                    string npmExe = @"C:\Program Files\nodejs\npm.cmd";
                    if (!File.Exists(npmExe)) npmExe = "npm.cmd";
                    try
                    {
                        var npmProc = Process.Start(new ProcessStartInfo
                        {
                            FileName = npmExe,
                            Arguments = "install",
                            WorkingDirectory = serviceDir,
                            UseShellExecute = false,
                            CreateNoWindow = true
                        });
                        npmProc?.WaitForExit(120000);
                    }
                    catch (Exception npmEx)
                    {
                        _logger.LogWarning(npmEx, "Failed running npm install in {Dir}", serviceDir);
                    }
                }

                string nodeExe = @"C:\Program Files\nodejs\node.exe";
                if (!File.Exists(nodeExe))
                {
                    nodeExe = "node";
                }

                var startInfo = new ProcessStartInfo
                {
                    FileName = nodeExe,
                    Arguments = "server.js",
                    WorkingDirectory = serviceDir,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                _nodeProcess = new Process { StartInfo = startInfo };
                _nodeProcess.OutputDataReceived += (s, e) =>
                {
                    if (!string.IsNullOrWhiteSpace(e.Data))
                        _logger.LogInformation("[WhatsApp Service] {Msg}", e.Data);
                };
                _nodeProcess.ErrorDataReceived += (s, e) =>
                {
                    if (!string.IsNullOrWhiteSpace(e.Data))
                        _logger.LogWarning("[WhatsApp Service Error] {Msg}", e.Data);
                };

                _nodeProcess.Start();
                _nodeProcess.BeginOutputReadLine();
                _nodeProcess.BeginErrorReadLine();
                _logger.LogInformation("Successfully spawned WhatsApp microservice (PID: {Pid}) in {Dir}", _nodeProcess.Id, serviceDir);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to launch WhatsApp node microservice");
            }
        }

        private void CleanupProcess()
        {
            try
            {
                if (_nodeProcess != null && !_nodeProcess.HasExited)
                {
                    _logger.LogInformation("Stopping WhatsApp Node microservice (PID: {Pid})...", _nodeProcess.Id);
                    _nodeProcess.Kill(true);
                    _nodeProcess.Dispose();
                    _nodeProcess = null;
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Error stopping WhatsApp node process");
            }
        }

        public override void Dispose()
        {
            CleanupProcess();
            base.Dispose();
        }
    }
}
