using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    public class HolidayModel
    {
        public DateTime DT { get; set; } = DateTime.Today;
        public string Description { get; set; } = string.Empty;

        public string DayName => DT.ToString("dddd");
        public string DateFormatted => DT.ToString("dd-MMM-yyyy");
    }
}
