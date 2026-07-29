using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    public class GatePassRowModel
    {
        public int EntryID { get; set; }           // 0 if new gate pass record
        public string EmpID { get; set; } = string.Empty;
        public string EmpName { get; set; } = string.Empty;
        public bool IsSelected { get; set; }
        public string OutTime { get; set; } = string.Empty; // Time1 (HH:mm)
        public string InTime { get; set; } = string.Empty;  // Time2 (HH:mm)
        public double GpHrs { get; set; }                   // Duration in hours
        public bool IsOfficial { get; set; }               // true = Official, false = Personal
        public string Remarks { get; set; } = string.Empty;

        public string DurationDisplay
        {
            get
            {
                if (GpHrs <= 0) return "-";
                int totalMins = (int)Math.Round(GpHrs * 60.0);
                int hrs = totalMins / 60;
                int mins = totalMins % 60;
                if (hrs > 0 && mins > 0) return $"{hrs}h {mins}m";
                if (hrs > 0) return $"{hrs}h";
                return $"{mins}m";
            }
        }
    }

    public class GatePassSaveDto
    {
        public DateTime Date { get; set; } = DateTime.Today;
        public List<GatePassRowModel> Rows { get; set; } = new();
    }
}
