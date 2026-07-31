using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    /// <summary>
    /// DTO passed from the Blazor page to the service/DAC layer when updating
    /// the monthly deduction installment amount for an employee's Long Term loan.
    /// </summary>
    public class AdjustDedAmtDto
    {
        /// <summary>Employee ID</summary>
        public string EmpID { get; set; } = string.Empty;

        /// <summary>Primary Key EntryID of the Advances row to update</summary>
        public long EntryID { get; set; }

        /// <summary>New monthly deduction amount</summary>
        public decimal NewDAmount { get; set; }
    }

    /// <summary>
    /// Summary model displaying active Long Term loan details for an employee.
    /// </summary>
    public class EmpDeductionInfoModel
    {
        public long EntryID { get; set; }
        public string EmpID { get; set; } = string.Empty;
        public decimal CurrentDAmount { get; set; }
        public decimal LoanAmount { get; set; }
        public decimal Balance { get; set; }
        public DateTime? IssueDate { get; set; }
        public string? Description { get; set; }
        public bool HasActiveLoan { get; set; }
    }
}
