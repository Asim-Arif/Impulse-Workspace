using System;

namespace DataAccessLibrary.Models.Setup
{
    public class RemoteAccessRequestModel
    {
        public int EntryID { get; set; }
        public int UserID { get; set; }
        public string UserName { get; set; } = string.Empty;
        public string ClientIP { get; set; } = string.Empty;
        public DateTime RequestDT { get; set; } = DateTime.Now;
        public string? Reason { get; set; }
        public string Status { get; set; } = "Pending"; // Pending, Approved, Rejected, Expired
        public int? DurationMinutes { get; set; }
        public DateTime? ExpiresAt { get; set; }
        public DateTime? ActionDT { get; set; }
        public int? ActionByUserID { get; set; }
        public string? ActionByUserName { get; set; }
        public string? Remarks { get; set; }

        // Additional Employee metadata for UI display
        public string? EmployeeName { get; set; }
        public string? DepartmentName { get; set; }
        public string? Designation { get; set; }
    }
}

