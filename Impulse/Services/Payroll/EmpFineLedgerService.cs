using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Services.Payroll
{
    public class EmpFineLedgerService : IEmpFineLedgerService
    {
        private readonly IEmpFineLedgerDataAccess _dataAccess;

        public EmpFineLedgerService(IEmpFineLedgerDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public async Task<(List<EmpFineLedgerItemModel> Items, EmpFineLedgerKpiSummaryModel Kpi)> GetFineLedgerDataAsync(EmpFineLedgerFilterDto filter)
        {
            decimal openBal = await _dataAccess.GetOpeningBalanceAsync(filter.EmpID, filter.FromDate);
            var dbItems     = await _dataAccess.GetFineLedgerAsync(filter);

            var fullList = new List<EmpFineLedgerItemModel>();

            // Insert Opening Balance Row
            fullList.Add(new EmpFineLedgerItemModel
            {
                EntryID          = 0,
                EmpID            = filter.EmpID != "0" ? filter.EmpID : string.Empty,
                EmpName          = filter.EmpID != "0" ? "Selected Employee" : "All Employees",
                DT               = filter.FromDate,
                Description      = "Previous Balance (Opening)",
                Amount           = openBal,
                RunningBalance   = openBal,
                IsOpeningBalance = true
            });

            decimal runningBal = openBal;
            decimal totalFines = 0;
            decimal totalCleared = 0;

            foreach (var item in dbItems)
            {
                if (item.Amount >= 0)
                    totalFines += item.Amount;
                else
                    totalCleared += System.Math.Abs(item.Amount);

                runningBal += item.Amount;
                item.RunningBalance = runningBal;
                fullList.Add(item);
            }

            var kpi = new EmpFineLedgerKpiSummaryModel
            {
                OpeningBalance   = openBal,
                TotalFinePosted  = totalFines,
                TotalFineCleared = totalCleared,
                ClosingBalance   = runningBal
            };

            return (fullList, kpi);
        }

        public async Task DeleteFineAsync(long entryId)
        {
            await _dataAccess.DeleteFineAsync(entryId);
        }
    }
}
