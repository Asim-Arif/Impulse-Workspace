using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    public class EmpTimeDetailRow
    {
        public int EntryID { get; set; }
        public string EmpID { get; set; } = string.Empty;
        public DateTime DT { get; set; }
        public string InTime { get; set; } = string.Empty;
        public string OutTime { get; set; } = string.Empty;
        public double Hrs { get; set; }
        public double PayableHrs { get; set; }
        public bool OverTime { get; set; }
        public string UserName { get; set; } = string.Empty;
    }

    public class ManualAttendanceInputDto
    {
        public string EmpID { get; set; } = string.Empty;
        public DateTime DT { get; set; } = DateTime.Today;
        public bool UseServerTime { get; set; } = false;
        public string InTime { get; set; } = "08:00";
        public string BrkOut { get; set; } = "13:00";
        public string BrkIn { get; set; } = "14:00";
        public string ChkOut { get; set; } = "17:00";
        public bool EnableOT { get; set; } = false;
        public string OTIn { get; set; } = "17:30";
        public string OTOut { get; set; } = "20:00";
        public string UserName { get; set; } = "System";
        public string ComputerName { get; set; } = "Unknown";
    }

    public class EmpShiftDefaultTimings
    {
        public string EnterTime { get; set; } = "08:00";
        public string ExitTime { get; set; } = "17:00";
        public string LunchInTime { get; set; } = "13:00";
        public string LunchOutTime { get; set; } = "14:00";
    }
}
