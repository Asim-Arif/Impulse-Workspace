using DataAccessLibrary.Models.ViewModels.Company;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Company
{
    public interface ICurrencyExchangeRatesService
    {
        Task<List<CurrencyExchangeRateModel>> GetRatesForDateAsync(DateTime date);
        Task<int> SaveRateAsync(DateTime date, string currency, float rate, int entryId);
    }
}
