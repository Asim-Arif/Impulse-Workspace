using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    public class MonthlyAttendanceDayRow
    {
        public int DayNo { get; set; }
        public DateTime Date { get; set; }
        public string DayName { get; set; } = string.Empty;
        public bool IsSundayOrHoliday { get; set; }
        public string HolidayName { get; set; } = string.Empty;
        public string Status { get; set; } = "Absent"; // Present, Leave, Absent
        public string InTime { get; set; } = string.Empty;  // HH:mm (24-hr format)
        public string OutTime { get; set; } = string.Empty; // HH:mm (24-hr format)
        public double OtHours { get; set; }
        public bool IsModified { get; set; }
    }

    public class MonthlyAttendanceSaveDto
    {
        public string EmpID { get; set; } = string.Empty;
        public int Year { get; set; }
        public int Month { get; set; }
        public List<MonthlyAttendanceDayRow> Rows { get; set; } = new();
    }
}
