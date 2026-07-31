using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public interface ISalarySheetService
    {
        /// <summary>
        /// Gets salary sheet records (either saved finalized or calculated unfinalized) along with KPI summary metrics.
        /// </summary>
        Task<(List<SalarySheetRowModel> Rows, SalarySheetKpiSummaryModel Kpi, bool IsFinalized)> GetSalarySheetDataAsync(SalarySheetFilterDto filter);

        /// <summary>
        /// Finalizes the salary sheet for the month/year/mode and posts Journal Vouchers.
        /// </summary>
        Task FinalizeSalarySheetAsync(int year, int month, bool isExternal, List<SalarySheetRowModel> rows);

        /// <summary>
        /// Updates the Salary_Paid flag for an individual employee's monthly salary.
        /// </summary>
        Task UpdateSalaryPaidStatusAsync(string empId, int year, int month, bool isPaid);

        /// <summary>
        /// Updates or inserts an overridden Long Term deduction amount for the month.
        /// </summary>
        Task UpdateLongTermDeductionOverrideAsync(string empId, DateTime monthDate, decimal amount);

        /// <summary>
        /// Adjusts negative net salary against open short term advances.
        /// </summary>
        Task TransferNegativeSalaryAsync(string empId, decimal amountToAdjust, DateTime monthDate);

        /// <summary>
        /// Posts centralized Journal Vouchers (JVs) for the finalized salary sheet of the specified year and month.
        /// Returns the number of department vouchers posted.
        /// </summary>
        Task<int> AddSalaryVouchersAsync(int year, int month);
    }
}
