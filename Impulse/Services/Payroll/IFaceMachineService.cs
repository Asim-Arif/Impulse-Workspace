using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    /// <summary>
    /// Communicates with ZKTeco IFace machines via COM interop using reflection with ParameterModifier
    /// for proper [out]/by-ref argument marshaling from SSR_GetGeneralLogData.
    /// Requires 64-bit zkemkeeper.dll (and supporting ZKTeco SDK DLLs) registered in System32.
    /// 
    /// Employee ID mapping: dynamic Company prefix + "-" + enrollNumber.PadLeft(5,'0').
    /// Punch time rule: if time <= 05:00, subtract 1 day (night-shift rollover).
    /// </summary>
    public class IFaceMachineService : IIFaceMachineService
    {
        private readonly ITakeAttendanceDataAccess _takeAttendance;
        private readonly IEmployeeDataAccess _employeeDataAccess;
        private readonly IBiometricAgentCoordinator _agentCoordinator;
        private readonly IConfiguration _config;
        private readonly ILogger<IFaceMachineService> _logger;

        public IFaceMachineService(
            ITakeAttendanceDataAccess takeAttendance,
            IEmployeeDataAccess employeeDataAccess,
            IBiometricAgentCoordinator agentCoordinator,
            IConfiguration config,
            ILogger<IFaceMachineService> logger)
        {
            _takeAttendance = takeAttendance;
            _employeeDataAccess = employeeDataAccess;
            _agentCoordinator = agentCoordinator;
            _config = config;
            _logger = logger;
        }

        public bool IsAgentConnected(string clientKey = "Default")
        {
            return _agentCoordinator.HasConnectedAgent(clientKey);
        }

        public async Task<(int recordsRead, int recordsProcessed, string message)> ReadAndProcessAsync(
            string ipAddress, int machineNo, int portNo = 4370, string userName = "System", string computerName = "Unknown", bool? clearLogs = null)
        {
            int recordsRead = 0;
            int recordsProcessed = 0;

            bool shouldClear = clearLogs ?? _config.GetValue<bool>("Biometric:ClearLogsAfterRead", false);

            string prefix = "EMP";
            try
            {
                prefix = await _employeeDataAccess.GetEmployeePrefixAsync();
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to resolve employee prefix from company; defaulting to EMP");
            }

            List<IFacePunchRecord> punches;

            // ── Auto-Routing: Cloud VPS (Agent) vs Local Server (Direct COM) ──
            if (_agentCoordinator.HasConnectedAgent())
            {
                _logger.LogInformation("Routing IFace machine read via connected Biometric Sync Agent for IP {IP} (ClearLogs={Clear})", ipAddress, shouldClear);
                var remoteResult = await _agentCoordinator.RequestPunchesAsync(ipAddress, machineNo, portNo, prefix, "Default", shouldClear);
                if (!remoteResult.success)
                {
                    return (0, 0, remoteResult.error);
                }
                punches = remoteResult.records;
                recordsRead = punches.Count;
            }
            else
            {
                // Local COM driver mode (Server is on-premises with ZKTeco SDK registered)
                Type? czKemType = Type.GetTypeFromProgID("zkemkeeper.ZKEM")
                               ?? Type.GetTypeFromProgID("zkemkeeper.CZKEM")
                               ?? Type.GetTypeFromCLSID(new Guid("00853A19-BD51-419B-9269-2DABE57EB61F"));
                if (czKemType == null)
                {
                    return (0, 0, "No office Biometric Sync Agent is connected, and ZKTeco COM component is not registered on this server. " +
                                  "Please start the office Sync Agent or register zkemkeeper.dll in C:\\Windows\\System32.");
                }

                try
                {
                    var result = await Task.Run(() =>
                    {
                        return ReadFromMachine(czKemType, ipAddress, machineNo, portNo, prefix, shouldClear);
                    });

                    if (!result.success)
                    {
                        return (0, 0, result.errorMessage);
                    }

                    punches = result.list;
                    recordsRead = result.totalRead;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "IFace machine read failed: {IP}", ipAddress);
                    return (0, 0, $"Failed to connect or read from machine ({ipAddress}): {ex.Message}");
                }
            }

            if (recordsRead == 0)
            {
                return (0, 0, $"No new punch records found on machine ({ipAddress}).");
            }

            // Process each punch through the central attendance engine
            foreach (var punch in punches)
            {
                try
                {
                    await _takeAttendance.ProcessPunchAsync(punch.EmpID, punch.AttDate, punch.AttTime, userName, computerName);
                    recordsProcessed++;
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Punch processing failed for EmpID={EmpID} DT={DT} Time={Time}",
                        punch.EmpID, punch.AttDate, punch.AttTime);
                }
            }

            string message = $"Read {recordsRead} punch records, processed {recordsProcessed} successfully. (Logs preserved on machine for trial)";
            return (recordsRead, recordsProcessed, message);
        }

        /// <summary>
        /// Synchronously connects to the ZKTeco IFace machine via COM reflection,
        /// reads all pending general log data, disconnects, and returns punch records.
        /// NOTE: ClearGLog is deliberately omitted during client trial runs to preserve machine logs.
        /// </summary>
        private (bool success, string errorMessage, List<IFacePunchRecord> list, int totalRead) ReadFromMachine(
            Type czKemType, string ipAddress, int machineNo, int portNo, string prefix = "EMP", bool clearLogs = false)
        {
            var records = new List<IFacePunchRecord>();
            int totalRead = 0;

            object? czkemInstance = null;
            try
            {
                czkemInstance = Activator.CreateInstance(czKemType);
            }
            catch (Exception ex)
            {
                return (false, $"Failed to initialize ZKTeco COM driver: {ex.Message}", records, 0);
            }

            if (czkemInstance == null)
            {
                return (false, "Failed to instantiate zkemkeeper.CZKEM COM object.", records, 0);
            }

            try
            {
                // Connect to machine
                object? connRet = czKemType.InvokeMember(
                    "Connect_Net",
                    BindingFlags.InvokeMethod,
                    null,
                    czkemInstance,
                    new object[] { ipAddress, portNo });

                bool connected = ToBool(connRet);
                if (!connected)
                {
                    int errCode = GetLastErrorCode(czKemType, czkemInstance);
                    string codeSuffix = errCode != 0 ? $" (Error Code: {errCode})" : "";
                    return (false, $"Could not connect to IFace machine at {ipAddress}:{portNo}{codeSuffix}. Please verify the machine IP and network connection.", records, 0);
                }

                try
                {
                    // Read general log data into machine's internal memory buffer
                    object? dataReadRet = czKemType.InvokeMember(
                        "ReadGeneralLogData",
                        BindingFlags.InvokeMethod,
                        null,
                        czkemInstance,
                        new object[] { machineNo });

                    bool dataRead = ToBool(dataReadRet);
                    if (!dataRead)
                    {
                        _logger.LogWarning("ReadGeneralLogData returned false for machine {No} at {IP}", machineNo, ipAddress);
                        return (true, string.Empty, records, 0);
                    }

                    // Prepare ParameterModifier for SSR_GetGeneralLogData:
                    // Param 0: dwMachineNumber (in)
                    // Param 1-10: dwEnrollNumber, dwVerifyMode, dwInOutMode, dwYear, dwMonth, dwDay, dwHour, dwMinute, dwSecond, dwWorkCode (all [out] / by-ref)
                    var paramMods = new ParameterModifier(11);
                    paramMods[0] = false;
                    for (int i = 1; i <= 10; i++) paramMods[i] = true;
                    var modifiers = new ParameterModifier[] { paramMods };

                    object[] logArgs = new object[11];
                    logArgs[0] = machineNo;
                    logArgs[1] = string.Empty;
                    for (int i = 2; i <= 10; i++) logArgs[i] = 0;

                    while (true)
                    {
                        object? getRet = czKemType.InvokeMember(
                            "SSR_GetGeneralLogData",
                            BindingFlags.InvokeMethod,
                            null,
                            czkemInstance,
                            logArgs,
                            modifiers,
                            null,
                            null);

                        if (!ToBool(getRet)) break;

                        totalRead++;

                        string enrollCode = logArgs[1]?.ToString() ?? string.Empty;
                        if (!int.TryParse(enrollCode, out int enrollNumber)) continue;

                        int year   = Convert.ToInt32(logArgs[4]);
                        int month  = Convert.ToInt32(logArgs[5]);
                        int day    = Convert.ToInt32(logArgs[6]);
                        int hour   = Convert.ToInt32(logArgs[7]);
                        int minute = Convert.ToInt32(logArgs[8]);

                        var attDate = new DateTime(year, month, day);
                        string attTime = $"{hour:D2}:{minute:D2}";

                        // Night-shift rollover: if punch is at or before 05:00, it belongs to previous day
                        bool isNextDay = (hour < 5) || (hour == 5 && minute == 0);
                        if (isNextDay) attDate = attDate.AddDays(-1);

                        // Dynamic EmpID format: $"{prefix}-00001"
                        string empId = $"{prefix}-{enrollNumber:D5}";

                        records.Add(new IFacePunchRecord
                        {
                            EmpID   = empId,
                            AttDate = attDate,
                            AttTime = attTime
                        });
                    }

                    // Clear logs from machine if requested
                    if (clearLogs)
                    {
                        try
                        {
                            czKemType.InvokeMember(
                                "ClearGLog",
                                BindingFlags.InvokeMethod,
                                null,
                                czkemInstance,
                                new object[] { machineNo });
                            _logger.LogInformation("Successfully cleared machine logs for machine {No} at {IP}", machineNo, ipAddress);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "Failed to clear log for machine {No}", machineNo);
                        }
                    }

                    try
                    {
                        czKemType.InvokeMember(
                            "EnableDevice",
                            BindingFlags.InvokeMethod,
                            null,
                            czkemInstance,
                            new object[] { machineNo, true });
                    }
                    catch { }
                }
                finally
                {
                    // Disconnect from machine
                    try
                    {
                        czKemType.InvokeMember(
                            "Disconnect",
                            BindingFlags.InvokeMethod,
                            null,
                            czkemInstance,
                            null);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Failed disconnecting from machine {IP}", ipAddress);
                    }
                }
            }
            finally
            {
                // Release COM Runtime Callable Wrapper (RCW) immediately
                if (czkemInstance != null && Marshal.IsComObject(czkemInstance))
                {
                    Marshal.FinalReleaseComObject(czkemInstance);
                }
            }

            return (true, string.Empty, records, totalRead);
        }

        private static bool ToBool(object? val)
        {
            if (val == null) return false;
            if (val is bool b) return b;
            if (int.TryParse(val.ToString(), out int i)) return i != 0;
            return false;
        }

        private static int GetLastErrorCode(Type czKemType, object czkemInstance)
        {
            try
            {
                var errMods = new ParameterModifier(1);
                errMods[0] = true;
                object[] args = new object[] { 0 };

                czKemType.InvokeMember(
                    "GetLastError",
                    BindingFlags.InvokeMethod,
                    null,
                    czkemInstance,
                    args,
                    new ParameterModifier[] { errMods },
                    null,
                    null);

                return Convert.ToInt32(args[0]);
            }
            catch
            {
                return 0;
            }
        }
    }
}
