using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    public class DepartmentListItemModel
    {
        public string DeptID { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public bool Active { get; set; }
        public int? Color { get; set; }
        public bool? TempDept { get; set; }
    }

    public class EmployeeListItemModel
    {
        public string EmpID { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string FName { get; set; } = string.Empty;
        public string Designation { get; set; } = string.Empty;
        public string Phone1 { get; set; } = string.Empty;
        public float StartingSalary { get; set; }
        public DateTime? JoinDate { get; set; }
        public bool Active { get; set; }
        public bool TempEmp { get; set; }
        public string DeptID { get; set; } = string.Empty;
        public string DeptName { get; set; } = string.Empty;
        public byte[]? EmpPic { get; set; }
    }

    public class EmployeeInActiveDetailsModel
    {
        public string EmpID { get; set; } = string.Empty;
        public DateTime DT { get; set; }
        public string InActiveRemarks { get; set; } = string.Empty;
        public string UserName { get; set; } = string.Empty;
        public string MachineName { get; set; } = string.Empty;
    }

    public class DepartmentModel
    {
        public string DeptID { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string? Description { get; set; }
        public bool Active { get; set; }
        public int? Color { get; set; }
        public string? Dept_ID { get; set; }
        public bool? TempDept { get; set; }
        public string? AccruedAccNo { get; set; }
        public string? SalaryAccNo { get; set; }
        public string? AccruedFoodAccNo { get; set; }
        public string? AccruedUnionAccNo { get; set; }
        public string? SchAccNo { get; set; }
        public string? AccruedEOBIAccNo { get; set; }
        public string? TaxAccNo { get; set; }
        public string? OnePercentAccNo { get; set; }
    }
}
