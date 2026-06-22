using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using CIP.Pages.Accounts;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Server.Kestrel.Core.Features;

namespace CIP.Services
{
    public class TrailBalance_BreakupService : ITrailBalance_BreakupService
    {
        private readonly ITrailBalance_BreakupDataAccess _dataAccess;
        private readonly IAuditService _auditService;
        private readonly IVoucherService _voucherService;
        private readonly IDBHelperService _dbhelperService;
        private readonly IAccountReportingAccess _accountreporting;

        public TrailBalance_BreakupService(ITrailBalance_BreakupDataAccess dataAccess, IAuditService auditService,
            IVoucherService voucherService, IDBHelperService dbhelperService, IAccountReportingAccess accountreporting)
        {
            _dataAccess = dataAccess;
            _auditService = auditService;
            _voucherService = voucherService;
            _dbhelperService = dbhelperService;
            _accountreporting = accountreporting;
        }
        public async Task<List<GenericDropDownModel>> GetValuesForDropDown(string strTableName, string strFieldName_ID, string strFieldName_Description, string strCondition = "")
        {
            return await _voucherService.GetValuesForDropDown(strTableName, strFieldName_ID, strFieldName_Description, strCondition);
        }
        public async Task<List<MainHeadsModel>> GetMainHeadsList()
        {
            return await _dataAccess.GetMainHeadsList();
        }
        //public async Task<List<TrailBalance_BreakupModel>> GenerateReport(DateTime fromDate, DateTime toDate, bool isSummary)
        //{
        //    var accounts = await _dataAccess.GetAccountsList(fromDate, toDate, isSummary);

        //    //foreach (var acc in accounts)
        //    //{
        //    //    var balances = await _dataAccess.CalculateAccountBalance(acc.AccNo, fromDate, toDate, acc.Parent);
        //    //
        //    //    acc.OpeningDr = balances.OpeningDr;
        //    //    acc.OpeningCr = balances.OpeningCr;
        //    //    acc.Debit = balances.Debit;
        //    //    acc.Credit = balances.Credit;
        //    //    acc.AdjDr = balances.AdjDr;
        //    //    acc.AdjCr = balances.AdjCr;
        //    //    acc.ClosingDr = balances.ClosingDr;
        //    //    acc.ClosingCr = balances.ClosingCr;
        //    //
        //    //}
        //    string accountCode = "";
        //    // Use GetAccountsList to obtain balances for all accounts (accountCode empty -> all)
        //    var allBalances = await _dataAccess.CalculateAccountBalance(accountCode, fromDate, toDate,false);

        //    var balanceMap = allBalances.ToDictionary(x => x.AccNo);
        //    foreach (var acc in accounts)
        //    {
        //        if (balanceMap.TryGetValue(acc.AccNo, out var bal))
        //        {
        //            acc.Debit = bal.Debit;
        //            acc.Credit = bal.Credit;
        //            acc.AdjDr = bal.AdjDr;
        //            acc.AdjCr = bal.AdjCr;
        //            // Calculate Closing locally
        //            double netClose = (acc.OpeningDr - acc.OpeningCr) + acc.Debit - acc.Credit;
        //            if (netClose >= 0) acc.ClosingDr = netClose; else acc.ClosingCr = Math.Abs(netClose);
        //        }
        //    }

        //    return accounts;
        //}

        public async Task<List<TrailBalance_BreakupModel>> GenerateReport(DateTime fromDate, DateTime toDate, bool isSummary)
        {
            var accounts = await _dataAccess.GetAccountsList(fromDate, toDate, isSummary);

            var allBalances = await _dataAccess.CalculateAccountBalances("", fromDate, toDate);

            var balanceLookup = allBalances.ToLookup(x => x.AccNo);

            foreach (var acc in accounts)
            {
                var bal = balanceLookup[acc.AccNo].FirstOrDefault();

                if (bal != null)
                {
                    acc.Debit = bal.Debit;
                    acc.Credit = bal.Credit;
                    acc.AdjDr = bal.AdjDr;
                    acc.AdjCr = bal.AdjCr;

                    acc.OpeningDr = bal.OpeningBalanceDr;
                    acc.OpeningCr = bal.OpeningBalanceCr;

                    double netClose = (acc.OpeningDr - acc.OpeningCr) + (acc.Debit - acc.Credit); //+ (acc.AdjDr - acc.AdjCr);

                    if (netClose >= 0)
                    {
                        acc.ClosingDr = netClose;
                        acc.ClosingCr = 0;
                    }
                    else
                    {
                        acc.ClosingDr = 0;
                        acc.ClosingCr = Math.Abs(netClose);
                    }

                    netClose = netClose + (acc.AdjDr - acc.AdjCr);
                    if (netClose >= 0)
                    {
                        acc.FinalClosingDr = netClose;
                        acc.FinalClosingCr = 0;
                    }
                    else
                    {
                        acc.FinalClosingDr = 0;
                        acc.FinalClosingCr = Math.Abs(netClose);
                    }



                }
            }

            return accounts;
        }

        public async Task<List<TrailBalance_BreakupModel>> GenerateAndSaveReport(List<TrailBalance_BreakupModel> myReportList,bool bWith_Opening_Closing)
        {
            if (myReportList != null && myReportList.Any())
            {
                // Call DataAccess to perform the bulk insert
                await _dataAccess.SaveLoadedDataToTempTable(myReportList, bWith_Opening_Closing);
            }
            return myReportList;
        }

    }
}