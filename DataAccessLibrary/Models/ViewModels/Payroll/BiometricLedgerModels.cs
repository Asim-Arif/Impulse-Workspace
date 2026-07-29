using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    /// <summary>
    /// Result row from SP_DeptSettings — one per department for the selected date.
    /// </summary>
    public class BiometricDeptSummaryRow
    {
        private string _deptId = string.Empty;
        public string DeptID { get => _deptId; set => _deptId = value?.Trim() ?? string.Empty; }
        public string Name         { get; set; } = string.Empty;
        public int    TotalEmps    { get; set; }
        public string EnterTime    { get; set; } = string.Empty;
        public string ExitTime     { get; set; } = string.Empty;
        public double WorkingHrs   { get; set; }
        public int    Color        { get; set; }          // VB long color value
        public int    TotalAbs     { get; set; }
        public int    LateComing   { get; set; }
        public int    ShortTimes   { get; set; }
        public int    ExcessTimes  { get; set; }

        // Computed helpers
        public int TotalPresent => TotalEmps - TotalAbs;

        /// <summary>
        /// Converts the legacy VB color int (BGR) to an HTML hex color string (#RRGGBB).
        /// </summary>
        public string HtmlColor
        {
            get
            {
                if (Color == 0) return "#343a40";
                int r = Color & 0xFF;
                int g = (Color >> 8) & 0xFF;
                int b = (Color >> 16) & 0xFF;
                return $"#{r:X2}{g:X2}{b:X2}";
            }
        }

        /// <summary>
        /// A very light (10% opacity) version of the dept color for row backgrounds.
        /// </summary>
        public string HtmlColorLight
        {
            get
            {
                if (Color == 0) return "rgba(52,58,64,0.08)";
                int r = Color & 0xFF;
                int g = (Color >> 8) & 0xFF;
                int b = (Color >> 16) & 0xFF;
                return $"rgba({r},{g},{b},0.10)";
            }
        }
    }

    /// <summary>
    /// One attendance row per employee per date — from VEmpTimes1 RIGHT OUTER JOIN VEmp.
    /// </summary>
    public class BiometricEmployeeRow
    {
        private string _deptId = string.Empty;
        private string _empId = string.Empty;
        public string DeptID { get => _deptId; set => _deptId = value?.Trim() ?? string.Empty; }
        public string EmpID  { get => _empId; set => _empId = value?.Trim() ?? string.Empty; }
        public string Name      { get; set; } = string.Empty;
        public int    EntryID   { get; set; }           // 0 = absent (no punch record)
        public DateTime? DT     { get; set; }
        public string InTime    { get; set; } = string.Empty;
        public string OutTime   { get; set; } = string.Empty;
        public double Hrs       { get; set; }
        public double LateHours { get; set; }
        public double ExtraHrs  { get; set; }

        public bool IsAbsent => EntryID == 0;

        /// <summary>Formats Hrs as H:MM string.</summary>
        public string NetTimeDisplay
        {
            get
            {
                if (Hrs <= 0) return string.Empty;
                int h = (int)Hrs;
                int m = (int)Math.Round((Hrs - h) * 60);
                return $"{h}:{m:D2}";
            }
        }

        /// <summary>Formats LateHours as H:MM string.</summary>
        public string LateTimeDisplay
        {
            get
            {
                if (LateHours <= 0) return string.Empty;
                int h = (int)LateHours;
                int m = (int)Math.Round((LateHours - h) * 60);
                return $"{h}:{m:D2}";
            }
        }
    }

    /// <summary>
    /// IFace (ZKTeco) biometric machine record from AttendanceMachine table.
    /// Only MachineType = 3 (IFace) machines are loaded.
    /// </summary>
    public class BiometricMachineItem
    {
        public int    EntryID     { get; set; }
        public string IPAddress   { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public int    MachineNo   { get; set; }
        public int    MachineType { get; set; }

        public string DisplayText => $"{Description} ({IPAddress})";
    }

    /// <summary>
    /// A punch record read from the IFace machine — used to call TakeAttendanceService.
    /// </summary>
    public class IFacePunchRecord
    {
        public string   EmpID    { get; set; } = string.Empty;
        public DateTime AttDate  { get; set; }
        public string   AttTime  { get; set; } = string.Empty;  // "HH:mm"
    }
}
