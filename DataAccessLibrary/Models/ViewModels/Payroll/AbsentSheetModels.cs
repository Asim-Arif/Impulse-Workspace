using System;
using System.Collections.Generic;
using System.Linq;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    public class AbsentSheetRowModel
    {
        public string EmpID { get; set; } = string.Empty;
        public string EmpName { get; set; } = string.Empty;
        public string DeptID { get; set; } = string.Empty;
        public string DeptName { get; set; } = string.Empty;
        public string Designation { get; set; } = string.Empty;
        public DateTime Date { get; set; }
        public string DayName => Date.ToString("dddd");
        public string FormattedDate => Date.ToString("dd-MMM-yyyy");
    }

    public class AbsentSheetFilterDto
    {
        public DateTime FromDate { get; set; } = DateTime.Today;
        public DateTime ToDate { get; set; } = DateTime.Today;
        public string? DeptID { get; set; }
        public int RangeOption { get; set; } = 0; // 0 = Today, 1 = Current Month, 2 = Custom
    }

    public class AbsentSheetResultModel
    {
        public List<AbsentSheetRowModel> Rows { get; set; } = new();
        public int TotalAbsents => Rows?.Count ?? 0;
    }
}
