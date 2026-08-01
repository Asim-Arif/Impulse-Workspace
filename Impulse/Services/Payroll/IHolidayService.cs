using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;

namespace Impulse.Services.Payroll
{
    public interface IHolidayService
    {
        Task<List<HolidayModel>> GetHolidaysByYearAsync(int year);
        Task<bool> AddHolidayAsync(HolidayModel model);
        Task<bool> UpdateHolidayAsync(DateTime dt, string description);
        Task<bool> DeleteHolidayAsync(DateTime dt);
    }
}
