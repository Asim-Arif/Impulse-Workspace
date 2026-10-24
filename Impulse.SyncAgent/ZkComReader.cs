using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.SyncAgent
{
    public static class ZkComReader
    {
        public static (bool Success, string Error, List<IFacePunchRecord> Punches) ReadPunches(
            string ipAddress, int machineNo, int portNo = 4370, string prefix = "EMP", bool clearLogs = false)
        {
            var records = new List<IFacePunchRecord>();

            Type? czKemType = Type.GetTypeFromProgID("zkemkeeper.ZKEM")
                           ?? Type.GetTypeFromProgID("zkemkeeper.CZKEM")
                           ?? Type.GetTypeFromCLSID(new Guid("00853A19-BD51-419B-9269-2DABE57EB61F"));

            if (czKemType == null)
            {
                return (false, "ZKTeco zkemkeeper COM component is not registered on this machine. Please register 64-bit zkemkeeper.dll in System32.", records);
            }

            object? czkemInstance = null;
            try
            {
                czkemInstance = Activator.CreateInstance(czKemType);
            }
            catch (Exception ex)
            {
                return (false, $"Failed to instantiate ZKTeco COM driver: {ex.Message}", records);
            }

            if (czkemInstance == null)
            {
                return (false, "Failed to instantiate zkemkeeper COM object.", records);
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

                if (!ToBool(connRet))
                {
                    int errCode = GetLastErrorCode(czKemType, czkemInstance);
                    string codeSuffix = errCode != 0 ? $" (Error Code: {errCode})" : "";
                    return (false, $"Could not connect to IFace machine at {ipAddress}:{portNo}{codeSuffix}. Please verify the machine IP and local network connection.", records);
                }

                try
                {
                    // Read general log data into machine buffer
                    object? dataReadRet = czKemType.InvokeMember(
                        "ReadGeneralLogData",
                        BindingFlags.InvokeMethod,
                        null,
                        czkemInstance,
                        new object[] { machineNo });

                    if (!ToBool(dataReadRet))
                    {
                        return (true, string.Empty, records); // No new logs
                    }

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

                        string enrollCode = logArgs[1]?.ToString() ?? string.Empty;
                        if (!int.TryParse(enrollCode, out int enrollNumber)) continue;

                        int year = Convert.ToInt32(logArgs[4]);
                        int month = Convert.ToInt32(logArgs[5]);
                        int day = Convert.ToInt32(logArgs[6]);
                        int hour = Convert.ToInt32(logArgs[7]);
                        int minute = Convert.ToInt32(logArgs[8]);

                        var attDate = new DateTime(year, month, day);
                        string attTime = $"{hour:D2}:{minute:D2}";

                        // Night-shift rollover
                        bool isNextDay = (hour < 5) || (hour == 5 && minute == 0);
                        if (isNextDay) attDate = attDate.AddDays(-1);

                        string empId = $"{prefix}-{enrollNumber:D5}";

                        records.Add(new IFacePunchRecord
                        {
                            EmpID = empId,
                            AttDate = attDate,
                            AttTime = attTime
                        });
                    }

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
                        }
                        catch { }
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
                    try
                    {
                        czKemType.InvokeMember(
                            "Disconnect",
                            BindingFlags.InvokeMethod,
                            null,
                            czkemInstance,
                            null);
                    }
                    catch { }
                }
            }
            finally
            {
                if (czkemInstance != null && Marshal.IsComObject(czkemInstance))
                {
                    Marshal.FinalReleaseComObject(czkemInstance);
                }
            }

            return (true, string.Empty, records);
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
