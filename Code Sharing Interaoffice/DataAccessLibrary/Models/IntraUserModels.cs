using System;

namespace DataAccessLibrary.Models.IntraOffice
{
    public class IntraDepartmentInfo
    {
        public string? Name { get; set; }
    }

    public class IntraDepartment
    {
        public string Id { get; set; } = "";
        public string Name { get; set; } = "";
    }

    public class IntraUserProfile
    {
        public int UserID { get; set; }
        public string UserName { get; set; } = string.Empty;
        public string Id { get => UserName; set => UserName = value ?? string.Empty; }
        public string? FullUserName { get; set; }
        public string FullName { get => FullUserName ?? EmployeeName ?? UserName; set => FullUserName = value; }
        public string? EmpID { get; set; }
        public string? EmployeeName { get; set; }
        public string? Designation { get; set; }
        public string? JobTitle { get => Designation; set => Designation = value; }
        public string? DepartmentId { get; set; }
        public string? DepartmentName { get; set; }
        public IntraDepartmentInfo? Department => !string.IsNullOrEmpty(DepartmentName) ? new IntraDepartmentInfo { Name = DepartmentName } : null;
        public string? Email { get; set; }
        public string? CellNo { get; set; }
        public string? PhoneNumber { get => CellNo; set => CellNo = value; }
        public string? Avatar { get; set; }
        public string? SignaturePath { get; set; }
        public UserStatus Status { get; set; } = UserStatus.Offline;
        public DateTime LastSeen { get; set; } = DateTime.UtcNow;
        public bool IsActive { get; set; } = true;
    }

    public class ApplicationUser : IntraUserProfile
    {
        public ApplicationUser() { }
        public ApplicationUser(IntraUserProfile p)
        {
            if (p != null)
            {
                UserID = p.UserID;
                UserName = p.UserName;
                FullUserName = p.FullUserName;
                EmpID = p.EmpID;
                EmployeeName = p.EmployeeName;
                Designation = p.Designation;
                DepartmentId = p.DepartmentId;
                DepartmentName = p.DepartmentName;
                Email = p.Email;
                CellNo = p.CellNo;
                Avatar = p.Avatar;
                Status = p.Status;
                LastSeen = p.LastSeen;
                IsActive = p.IsActive;
            }
        }
    }
}
