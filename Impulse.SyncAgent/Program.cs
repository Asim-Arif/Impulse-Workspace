using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Microsoft.AspNetCore.SignalR.Client;
using Microsoft.Extensions.Configuration;

namespace Impulse.SyncAgent
{
    internal class Program
    {
        private static async Task Main(string[] args)
        {
            Console.Title = "Impulse Biometric Sync Agent";
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("=============================================================");
            Console.WriteLine("           IMPULSE BIOMETRIC SYNC AGENT                      ");
            Console.WriteLine("   Bridges Cloud VPS with Local Office Attendance Machines    ");
            Console.WriteLine("=============================================================");
            Console.ResetColor();

            // Load configuration
            var config = new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            string serverUrl = config["SyncAgent:ServerUrl"] ?? "http://localhost:5000/hubs/biometric-agent";
            string clientKey = config["SyncAgent:ClientKey"] ?? "Default";
            string locationName = config["SyncAgent:LocationName"] ?? Environment.MachineName;

            Console.WriteLine($"Server URL:     {serverUrl}");
            Console.WriteLine($"Client Key:     {clientKey}");
            Console.WriteLine($"Office Station: {locationName}");
            Console.WriteLine("-------------------------------------------------------------");

            var connection = new HubConnectionBuilder()
                .WithUrl(serverUrl)
                .WithAutomaticReconnect(new[] { TimeSpan.Zero, TimeSpan.FromSeconds(2), TimeSpan.FromSeconds(5), TimeSpan.FromSeconds(10), TimeSpan.FromSeconds(30) })
                .Build();

            connection.Reconnecting += error =>
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Connection lost. Reconnecting... ({error?.Message})");
                Console.ResetColor();
                return Task.CompletedTask;
            };

            connection.Reconnected += async connectionId =>
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Reconnected! Re-registering with server...");
                Console.ResetColor();
                try
                {
                    await connection.SendAsync("RegisterAgent", clientKey, locationName);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to re-register: {ex.Message}");
                }
            };

            connection.Closed += async error =>
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Connection closed. Retrying in 5 seconds...");
                Console.ResetColor();
                await Task.Delay(5000);
                await StartConnectionWithRetryAsync(connection, clientKey, locationName);
            };

            // Handle read commands from VPS
            connection.On<BiometricAgentReadCommand>("ExecuteRead", async (cmd) =>
            {
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine($"\n[{DateTime.Now:HH:mm:ss}] [READ REQUEST] Target Machine: {cmd.IPAddress}:{cmd.PortNo} (MachineNo={cmd.MachineNo})");
                Console.ResetColor();

                var (success, error, punches) = await Task.Run(() =>
                    ZkComReader.ReadPunches(cmd.IPAddress, cmd.MachineNo, cmd.PortNo, cmd.Prefix, cmd.ClearLogs));

                if (success)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] [SUCCESS] Read {punches.Count} punches from {cmd.IPAddress}. Sending to server...");
                    Console.ResetColor();
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] [FAILED] {error}");
                    Console.ResetColor();
                }

                try
                {
                    await connection.SendAsync("ReportReadResult", new BiometricAgentReadResult
                    {
                        RequestId = cmd.RequestId,
                        Success = success,
                        ErrorMessage = error,
                        Punches = punches
                    });
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Results acknowledged by server.");
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Failed to report results to server: {ex.Message}");
                    Console.ResetColor();
                }
            });

            // Start connection
            await StartConnectionWithRetryAsync(connection, clientKey, locationName);

            Console.WriteLine("\n[STATUS] Agent is ready and listening. You can minimize this window.");
            Console.WriteLine("Press Ctrl+C to stop.\n");

            var cts = new CancellationTokenSource();
            Console.CancelKeyPress += (s, e) =>
            {
                e.Cancel = true;
                cts.Cancel();
            };

            try
            {
                await Task.Delay(Timeout.Infinite, cts.Token);
            }
            catch (TaskCanceledException)
            {
                Console.WriteLine("Shutting down agent...");
            }
        }

        private static async Task StartConnectionWithRetryAsync(HubConnection connection, string clientKey, string locationName)
        {
            while (true)
            {
                try
                {
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Connecting to Impulse server...");
                    await connection.StartAsync();
                    await connection.SendAsync("RegisterAgent", clientKey, locationName);
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] CONNECTED! Registered as '{clientKey}' ({locationName}).");
                    Console.ResetColor();
                    break;
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Connect attempt failed: {ex.Message}. Retrying in 5 seconds...");
                    Console.ResetColor();
                    await Task.Delay(5000);
                }
            }
        }
    }
}
