using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Budgeting;

namespace DataAccessLibrary.Interface.Budgeting
{
    public interface IBudgetDataAccess
    {
        List<FinancialYearOption> GetFinancialYears(int startYear = 2010);
        Task<List<MasterHeadOption>> GetMasterHeadsAsync(bool isExpenses);
        Task<List<MonthwiseBudgetRow>> GetMonthwiseBudgetAsync(string financialYear, string headCode = "0", bool isRevision = false);
        Task<bool> SaveMonthwiseCellAsync(string financialYear, string accNo, string monthColumn, decimal amount, bool isRevision = false);
        Task<bool> UpdateMonthwiseRevisionsAsync(string financialYear);
        Task<List<ExpenseBudgetRow>> GetExpenseBudgetAsync(string financialYear, string headCode = "0");
        Task<bool> SaveExpenseCellAsync(string financialYear, string accNo, string column, decimal amount);
        Task<bool> UpdateExpenseRevisionsAsync(string financialYear);
        Task<List<BudgetReportRow>> GetQuarterlyReportDataAsync(string financialYear, int quarter, string headCode, bool isExpenses);
        Task<List<BudgetReportRow>> GetDateRangeReportDataAsync(string financialYear, DateTime fromDate, DateTime toDate, string headCode, bool isExpenses);
        Task<List<BudgetReportRow>> GetBudgetAtGlanceDataAsync(string financialYear, string headCode);
    }
}

