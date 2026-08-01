using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.DAC.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public class HolidayService : IHolidayService
    {
        private readonly IHolidayDataAccess _dac;

        public HolidayService(IHolidayDataAccess dac)
        {
            _dac = dac;
        }

        public async Task<List<HolidayModel>> GetHolidaysByYearAsync(int year)
        {
            return await _dac.GetHolidaysByYearAsync(year);
        }

        public async Task<bool> AddHolidayAsync(HolidayModel model)
        {
            return await _dac.AddHolidayAsync(model);
        }

        public async Task<bool> UpdateHolidayAsync(DateTime dt, string description)
        {
            return await _dac.UpdateHolidayAsync(dt, description);
        }

        public async Task<bool> DeleteHolidayAsync(DateTime dt)
        {
            return await _dac.DeleteHolidayAsync(dt);
        }
    }
}
