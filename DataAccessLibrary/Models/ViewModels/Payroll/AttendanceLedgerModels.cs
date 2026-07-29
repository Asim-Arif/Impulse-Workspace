using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    public class AttendanceLedgerFilterModel
    {
        public string DeptID { get; set; } = "0";
        public string EmpID { get; set; } = string.Empty;
        public int Month { get; set; } = DateTime.Today.Month;
        public int Year { get; set; } = DateTime.Today.Year;
    }

    /// <summary>Row from VEmpLedger (AttendanceSheet + Leaves + EmpOTHrs) for one employee / month.</summary>
    public class AttendanceLedgerEntryModel
    {
        public DateTime DT { get; set; }
        public int AttDay { get; set; }
        public byte? Attendance { get; set; }
        public double CL { get; set; }
        public double SL { get; set; }
        public double AL { get; set; }
        public double CPL { get; set; }
        public double LWP { get; set; }
        public double ML { get; set; }
        public double SAL { get; set; }
        public double? OTHrsOther { get; set; }
    }

    /// <summary>Row from VEmpTimes2 (pivoted first/second/overtime punches) for one employee / day.</summary>
    public class AttendanceTimeModel
    {
        public DateTime DT { get; set; }
        public string FirstInTime { get; set; } = string.Empty;
        public string FirstOutTime { get; set; } = string.Empty;
        public string SecondInTime { get; set; } = string.Empty;
        public string SecondOutTime { get; set; } = string.Empty;
        public string OTInTime { get; set; } = string.Empty;
        public string OTOutTime { get; set; } = string.Empty;
        public double Hrs { get; set; }
        public double PayableHrs { get; set; }
        public double OTHrs { get; set; }
        public double LateHours { get; set; }
    }

    public class HolidayItemModel
    {
        public int Day { get; set; }
        public string Description { get; set; } = string.Empty;
    }

    /// <summary>One display row of the Attendance Ledger grid (one per day of the month).</summary>
    public class AttendanceLedgerDayRow
    {
        public DateTime Date { get; set; }
        public string DayName { get; set; } = string.Empty;
        public bool IsSunday { get; set; }
        public bool IsHoliday { get; set; }
        public bool IsOffDay => IsSunday || IsHoliday;
        public bool HasRecord { get; set; }

        /// <summary>"", "P", "A", "L" or "HL" — mirrors the legacy status column.</summary>
        public string Status { get; set; } = string.Empty;
        public string LeaveType { get; set; } = string.Empty;

        public string FirstInTime { get; set; } = string.Empty;
        public string FirstOutTime { get; set; } = string.Empty;
        public string SecondInTime { get; set; } = string.Empty;
        public string SecondOutTime { get; set; } = string.Empty;
        public string OTInTime { get; set; } = string.Empty;
        public string OTOutTime { get; set; } = string.Empty;

        public double Hrs { get; set; }
        public double PayableHrs { get; set; }
        public double OTHrs { get; set; }
        public double LateHours { get; set; }
        public double OTHrsOther { get; set; }

        public string ShortLeaves { get; set; } = string.Empty;

        public bool ShowTimes { get; set; }
    }

    /// <summary>C# port of the legacy PayrollFunctions.bas getHrsMin(): decimal hours -> "HH:MM".</summary>
    public static class AttendanceLedgerFormatter
    {
        public static string GetHrsMin(double hrs, bool showDashInsteadOfZero = false)
        {
            int totalMins = (int)Math.Round(hrs * 60, MidpointRounding.ToEven);
            int tHrs = totalMins / 60;
            int tMin = totalMins % 60;

            if (tHrs == 0 && tMin == 0)
                return showDashInsteadOfZero ? "-" : "0";

            return $"{tHrs:00}:{tMin:00}";
        }
    }
}
