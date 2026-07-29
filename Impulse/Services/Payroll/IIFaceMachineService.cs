using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    /// <summary>
    /// Handles communication with ZKTeco IFace biometric machines.
    /// Reads attendance punch records and routes each punch through the central TakeAttendance engine.
    /// </summary>
    public interface IIFaceMachineService
    {
        /// <summary>
        /// Connects to the IFace machine at the given IP, reads all pending attendance log data,
        /// processes each punch via TakeAttendanceService, clears the machine log, and disconnects.
        /// Returns a summary message describing how many records were processed.
        /// </summary>
        Task<(int recordsRead, int recordsProcessed, string message)> ReadAndProcessAsync(
            string ipAddress, int machineNo, int portNo = 4370);
    }
}
