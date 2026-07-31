using DataAccessLibrary.Models.ViewModels.Payroll;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Payroll
{
    public interface ISalarySheetDataAccess
    {
        /// <summary>
        /// Checks if salary sheet for the month/year/isExternal has already been finalized.
        /// </summary>
        Task<bool> IsSalarySheetFinalizedAsync(int year, int month, bool isExternal);

        /// <summary>
        /// Fetches saved finalized salary sheet records from VMonthlySalaries.
        /// </summary>
        Task<List<SalarySheetRowModel>> GetFinalizedSalarySheetAsync(SalarySheetFilterDto filter);

        /// <summary>
        /// Calculates unfinalized salary sheet records using VAdvanceSalary stored procedures and attendance rules.
        /// </summary>
        Task<List<SalarySheetRowModel>> GetUnfinalizedSalarySheetAsync(SalarySheetFilterDto filter);

        /// <summary>
        /// Finalizes the salary sheet atomically inside a database transaction:
        /// 1. Inserts into MonthlySalaries
        /// 2. Updates AmtCleared for Long Term loans
        /// 3. Updates Advances.Cleared for Short Term advances
        /// 4. Posts accounting Journal Vouchers (JVs) per department.
        /// </summary>
        Task FinalizeSalarySheetAsync(int year, int month, bool isExternal, List<SalarySheetRowModel> rows, string userName, string machineName);

        /// <summary>
        /// Updates the Salary_Paid flag for an individual employee's monthly salary.
        /// </summary>
        Task UpdateSalaryPaidStatusAsync(string empId, int year, int month, bool isPaid, string userName, string machineName);

        /// <summary>
        /// Updates or inserts an overridden Long Term deduction amount for the month.
        /// </summary>
        Task UpdateLongTermDeductionOverrideAsync(string empId, DateTime monthDate, decimal amount, string userName, string machineName);

        /// <summary>
        /// Adjusts negative net salary against open short term advances.
        /// </summary>
        Task TransferNegativeSalaryAsync(string empId, decimal amountToAdjust, DateTime monthDate);

        /// <summary>
        /// Fetches aggregated department salary totals and account mappings for JV posting.
        /// </summary>
        Task<List<DepartmentSalaryVoucherSummaryDto>> GetDepartmentSalaryVoucherSummariesAsync(int year, int month);

        /// <summary>
        /// Fetches default short term and long term advance account numbers from GeneralData.
        /// </summary>
        Task<(string ShortTermAccNo, string LongTermAccNo)> GetSalaryVoucherGeneralAccountsAsync();
    }
}
