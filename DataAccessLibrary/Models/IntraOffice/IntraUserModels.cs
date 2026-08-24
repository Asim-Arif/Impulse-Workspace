using System;

namespace DataAccessLibrary.Models.IntraOffice
{
    public class IntraUserProfile
    {
        public int UserID { get; set; }
        public string UserName { get; set; } = string.Empty;
        public string? FullUserName { get; set; }
        public string? EmpID { get; set; }
        public string? EmployeeName { get; set; }
        public string? Designation { get; set; }
        public string? DepartmentId { get; set; }
        public string? DepartmentName { get; set; }
        public string? Email { get; set; }
        public string? CellNo { get; set; }
        public string? Avatar { get; set; }
        public UserStatus Status { get; set; } = UserStatus.Offline;
        public DateTime LastSeen { get; set; } = DateTime.UtcNow;
        public bool IsActive { get; set; } = true;
    }
}
