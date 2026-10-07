using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    public enum AttendanceWorkflowActionType
    {
        ManualSave = 1,
        ManualDelete = 2,
        MonthlySave = 3,
        MonthlyClearDate = 4
    }

    public class AttendanceWorkflowRequestDto
    {
        public int TaskId { get; set; }
        public AttendanceWorkflowActionType ActionType { get; set; }
        public string EmpID { get; set; } = string.Empty;
        public string EmployeeName { get; set; } = string.Empty;
        public string DepartmentName { get; set; } = string.Empty;
        public int Year { get; set; }
        public int Month { get; set; }
        public DateTime? AttendanceDate { get; set; }
        public string Reason { get; set; } = string.Empty;
        public string OriginatorUserName { get; set; } = "System";
        public string MachineName { get; set; } = string.Empty;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        // Payloads
        public ManualAttendanceInputDto? ManualInput { get; set; }
        public MonthlyAttendanceSaveDto? MonthlyInput { get; set; }

        // Comparison Diffs for UI Display
        public List<AttendanceDiffRow> DiffRows { get; set; } = new();
    }

    public class AttendanceDiffRow
    {
        public DateTime Date { get; set; }
        public string DayName { get; set; } = string.Empty;
        
        // Previous (Current in DB)
        public string PrevInTime { get; set; } = string.Empty;
        public string PrevOutTime { get; set; } = string.Empty;
        public string PrevStatus { get; set; } = string.Empty;
        public double PrevOtHours { get; set; }

        // Proposed (New)
        public string NewInTime { get; set; } = string.Empty;
        public string NewOutTime { get; set; } = string.Empty;
        public string NewStatus { get; set; } = string.Empty;
        public double NewOtHours { get; set; }

        public bool IsDeleted { get; set; }
    }
}
