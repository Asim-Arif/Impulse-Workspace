using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public class GatePassService : IGatePassService
    {
        private readonly IGatePassDataAccess _dac;

        public GatePassService(IGatePassDataAccess dac)
        {
            _dac = dac;
        }

        public Task<List<GatePassRowModel>> GetGatePassListAsync(string deptId, DateTime date)
            => _dac.GetGatePassListAsync(deptId, date);

        public Task<bool> SaveGatePassesAsync(GatePassSaveDto input)
            => _dac.SaveGatePassesAsync(input);
    }
}
