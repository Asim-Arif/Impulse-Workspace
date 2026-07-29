using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public interface IGatePassService
    {
        Task<List<GatePassRowModel>> GetGatePassListAsync(string deptId, DateTime date);
        Task<bool> SaveGatePassesAsync(GatePassSaveDto input);
    }
}
