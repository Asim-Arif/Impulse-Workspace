using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface IGatePassDataAccess
    {
        /// <summary>
        /// Gets present employees and existing gate passes for a specific department and date.
        /// </summary>
        Task<List<GatePassRowModel>> GetGatePassListAsync(string deptId, DateTime date);

        /// <summary>
        /// Saves gate pass rows (inserts new, updates existing, deletes unchecked).
        /// </summary>
        Task<bool> SaveGatePassesAsync(GatePassSaveDto input);
    }
}
