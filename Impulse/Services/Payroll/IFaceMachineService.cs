using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    /// <summary>
    /// Communicates with ZKTeco IFace machines via COM interop (late-binding dynamic).
    /// Requires zkemkeeper.dll (or the ZKTeco SDK) to be registered on the server.
    /// 
    /// Employee ID mapping: "EMR-" + enrollNumber.PadLeft(5,'0')  (same as legacy).
    /// Punch time rule: if time <= 05:00, subtract 1 day (night-shift rollover).
    /// </summary>
    public class IFaceMachineService : IIFaceMachineService
    {
        private readonly ITakeAttendanceDataAccess _takeAttendance;
        private readonly ILogger<IFaceMachineService> _logger;

        public IFaceMachineService(
            ITakeAttendanceDataAccess takeAttendance,
            ILogger<IFaceMachineService> logger)
        {
            _takeAttendance = takeAttendance;
            _logger = logger;
        }

        public async Task<(int recordsRead, int recordsProcessed, string message)> ReadAndProcessAsync(
            string ipAddress, int machineNo, int portNo = 4370)
        {
            int recordsRead = 0;
            int recordsProcessed = 0;

            // Collect punch records first (COM work must be synchronous)
            List<IFacePunchRecord> punches;
            try
            {
                punches = ReadFromMachine(ipAddress, machineNo, portNo, out recordsRead);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "IFace machine read failed: {IP}", ipAddress);
                return (0, 0, $"Failed to connect or read from machine ({ipAddress}): {ex.Message}");
            }

            // Process each punch through the central attendance engine
            foreach (var punch in punches)
            {
                try
                {
                    await _takeAttendance.ProcessPunchAsync(punch.EmpID, punch.AttDate, punch.AttTime);
                    recordsProcessed++;
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Punch processing failed for EmpID={EmpID} DT={DT} Time={Time}",
                        punch.EmpID, punch.AttDate, punch.AttTime);
                }
            }

            string message = $"Read {recordsRead} punch records, processed {recordsProcessed} successfully.";
            return (recordsRead, recordsProcessed, message);
        }

        /// <summary>
        /// Synchronously connects to the ZKTeco IFace machine via COM (late-binding),
        /// reads all pending general log data, clears the log, disconnects, and returns punch records.
        /// </summary>
        private List<IFacePunchRecord> ReadFromMachine(string ipAddress, int machineNo, int portNo, out int totalRead)
        {
            totalRead = 0;
            var records = new List<IFacePunchRecord>();

            // Late-bind to zkemkeeper.CZKEM COM object
            Type? czKemType = Type.GetTypeFromProgID("zkemkeeper.CZKEM");
            if (czKemType == null)
                throw new InvalidOperationException("ZKTeco zkemkeeper COM component is not registered on this server. Please install the ZKTeco SDK.");

            dynamic czkem = Activator.CreateInstance(czKemType)!;

            // Connect
            bool connected = czkem.Connect_Net(ipAddress, portNo) == 1;
            if (!connected)
                throw new InvalidOperationException($"Could not connect to IFace machine at {ipAddress}:{portNo}. Check IP and network.");

            try
            {
                // Read general log data into machine's internal buffer
                bool dataRead = czkem.ReadGeneralLogData(machineNo) == 1;
                if (!dataRead)
                {
                    _logger.LogWarning("ReadGeneralLogData returned false for machine {No} at {IP}", machineNo, ipAddress);
                    return records;
                }

                // Iterate all buffered records
                string  enrollCode  = string.Empty;
                int     verifyMode  = 0;
                int     inOutMode   = 0;
                int     year = 0, month = 0, day = 0, hour = 0, minute = 0, second = 0, workerCode = 0;

                while (czkem.SSR_GetGeneralLogData(
                           machineNo,
                           out enrollCode,
                           out verifyMode,
                           out inOutMode,
                           out year, out month, out day,
                           out hour, out minute, out second,
                           out workerCode) != 0)
                {
                    totalRead++;

                    if (!int.TryParse(enrollCode, out int enrollNumber)) continue;

                    var attDate = new DateTime(year, month, day);
                    string attTime = $"{hour:D2}:{minute:D2}";

                    // Night-shift rollover: if punch is at or before 05:00, it belongs to previous day
                    bool isNextDay = (hour < 5) || (hour == 5 && minute == 0);
                    if (isNextDay) attDate = attDate.AddDays(-1);

                    // Legacy EmpID format: "EMR-00001"
                    string empId = $"EMR-{enrollNumber:D5}";

                    records.Add(new IFacePunchRecord
                    {
                        EmpID   = empId,
                        AttDate = attDate,
                        AttTime = attTime
                    });
                }

                // Clear log from machine after successful read
                czkem.ClearGLog(machineNo);
                czkem.EnableDevice(machineNo, true);
            }
            finally
            {
                czkem.Disconnect();
            }

            return records;
        }
    }
}
