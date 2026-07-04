using DataAccessLibrary.Interface.Company;
using DataAccessLibrary.Models.ViewModels.Company;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Company
{
    public class CurrencyExchangeRatesService : ICurrencyExchangeRatesService
    {
        private readonly ICurrencyExchangeRatesDataAccess _dataAccess;

        public CurrencyExchangeRatesService(ICurrencyExchangeRatesDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<CurrencyExchangeRateModel>> GetRatesForDateAsync(DateTime date)
        {
            return _dataAccess.GetRatesForDateAsync(date);
        }

        public Task<int> SaveRateAsync(DateTime date, string currency, float rate, int entryId)
        {
            return _dataAccess.SaveRateAsync(date, currency, rate, entryId);
        }
    }
}
