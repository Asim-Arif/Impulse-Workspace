using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    public class SalaryHistoryRowModel
    {
        public string DeptID { get; set; } = string.Empty;
        public string DeptName { get; set; } = string.Empty;
        public string EmpID { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string FName { get; set; } = string.Empty;
        public string Rel { get; set; } = string.Empty;
        public string Designation { get; set; } = string.Empty;
        public DateTime? DT { get; set; }
        public string Description { get; set; } = string.Empty;
        public double Salary { get; set; }
        public double SalaryExt { get; set; }
        public double? PreviousSalary { get; set; }
    }

    public class SalaryHistoryFilterModel
    {
        public string DeptID { get; set; } = "0";
        public string EmpID { get; set; } = "0";
        public DateTime FromDate { get; set; } = DateTime.Today.AddMonths(-1);
        public DateTime ToDate { get; set; } = DateTime.Today;
    }

    public class ChangeSalaryRequestModel
    {
        public string EmpID { get; set; } = string.Empty;
        public DateTime EffectiveDate { get; set; } = DateTime.Today;
        public double NewSalary { get; set; }
        public double NewSalaryExt { get; set; }
        public string Description { get; set; } = "Salary Adjustment";
        public string UserName { get; set; } = "System";
    }
}
