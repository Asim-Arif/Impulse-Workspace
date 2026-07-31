using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    /// <summary>
    /// DTO passed from the Blazor UI to save or update an employee hold salary record.
    /// Matches legacy VB6 <c>frmEmpHoldSalary.frm</c> fields.
    /// </summary>
    public class HoldSalaryDto
    {
        public string EmpID { get; set; } = string.Empty;
        public int Year { get; set; } = DateTime.Today.Year;
        public int Month { get; set; } = DateTime.Today.Month;
        public int HoldAmt { get; set; }
        public string Description { get; set; } = string.Empty;
    }

    /// <summary>
    /// Summary model containing the employee's available salary balance and existing hold amount
    /// for the selected month/year.
    /// </summary>
    public class HoldSalarySummaryModel
    {
        public string EmpID { get; set; } = string.Empty;
        public string EmpName { get; set; } = string.Empty;
        public string DeptID { get; set; } = string.Empty;
        public string DeptName { get; set; } = string.Empty;

        /// <summary>
        /// Calculated as: Balance + ISNULL(HoldSalaryAmt, 0) from MonthlySalaries for the month end date.
        /// </summary>
        public decimal AvailableBalance { get; set; }

        /// <summary>
        /// Existing hold amount from EmpHoldSalary table for the 1st of the month.
        /// </summary>
        public int ExistingHoldAmt { get; set; }
        public string ExistingDescription { get; set; } = string.Empty;

        public bool IsEligible { get; set; } = true;
        public string IneligibilityReason { get; set; } = string.Empty;
    }

    /// <summary>
    /// Represents an existing hold salary entry in the grid for the selected month/year.
    /// </summary>
    public class HoldSalaryListItemModel
    {
        public int EntryID { get; set; }
        public string EmpID { get; set; } = string.Empty;
        public string EmpName { get; set; } = string.Empty;
        public string DeptID { get; set; } = string.Empty;
        public string DeptName { get; set; } = string.Empty;
        public DateTime DT { get; set; }
        public int HoldAmt { get; set; }
        public string Description { get; set; } = string.Empty;
        public DateTime? DTEntry { get; set; }
        public string UserName { get; set; } = string.Empty;
        public string MachineName { get; set; } = string.Empty;
    }
}
