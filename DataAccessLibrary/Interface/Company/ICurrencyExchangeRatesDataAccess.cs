using DataAccessLibrary.Models.ViewModels.Company;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DataAccessLibrary.Interface.Company
{
    public interface ICurrencyExchangeRatesDataAccess
    {
        Task<List<CurrencyExchangeRateModel>> GetRatesForDateAsync(DateTime date);
        Task<int> SaveRateAsync(DateTime date, string currency, float rate, int entryId);
    }
}
