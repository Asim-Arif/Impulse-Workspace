using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    public class EmployeeLeaveInputDto
    {
        public string EmpID { get; set; } = string.Empty;
        public DateTime FromDate { get; set; } = DateTime.Today;
        public DateTime ToDate { get; set; } = DateTime.Today;
        public bool IsAbsent { get; set; } = false;             // false = Leave, true = Absent
        public string DurationMode { get; set; } = "FullDay";   // FullDay, HalfDay, ShortLeave
        public string FromTime { get; set; } = "08:00";
        public string ToTime { get; set; } = "17:00";
        public string LeaveTypeCode { get; set; } = "CL";
        public string LeaveTypeName { get; set; } = "Casual Leave";
        public string AuthPerson { get; set; } = string.Empty;
        public string Reason { get; set; } = string.Empty;
        public bool WithDeduction { get; set; } = false;
    }

    public class LeaveCategoryBalance
    {
        public double Available { get; set; }
        public double Used { get; set; }
        public double ThisMonth { get; set; }
        public double Balance => Available - Used - ThisMonth;
    }

    public class LeaveBalanceMatrixDto
    {
        public LeaveCategoryBalance CL { get; set; } = new();
        public LeaveCategoryBalance SL { get; set; } = new();
        public LeaveCategoryBalance AL { get; set; } = new();
        public LeaveCategoryBalance SAL { get; set; } = new();
        public LeaveCategoryBalance LWP { get; set; } = new();
        public LeaveCategoryBalance HL { get; set; } = new();
    }

    public class LeaveTypeLookupModel
    {
        public string Code { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public double Limit { get; set; }
    }

    public class UserLookupModel
    {
        public string UserName { get; set; } = string.Empty;
    }
}
