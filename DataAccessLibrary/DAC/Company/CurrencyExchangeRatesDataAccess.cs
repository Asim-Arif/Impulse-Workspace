using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Company;
using DataAccessLibrary.Models.ViewModels.Company;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.Company
{
    public class CurrencyExchangeRatesDataAccess : ICurrencyExchangeRatesDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<CurrencyExchangeRatesDataAccess> _logger;

        public CurrencyExchangeRatesDataAccess(IConfiguration configuration, ILogger<CurrencyExchangeRatesDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        public async Task<List<CurrencyExchangeRateModel>> GetRatesForDateAsync(DateTime date)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                
                // 1. Get all distinct active currencies from ForeignCustomers
                var currencies = await db.QueryAsync<string>("SELECT DISTINCT Curr FROM ForeignCustomers WHERE ISNULL(Curr, '') <> ''");
                
                if (!currencies.Any())
                    return new List<CurrencyExchangeRateModel>();

                var resultList = new List<CurrencyExchangeRateModel>();
                var formattedDate = date.ToString("yyyy-MM-dd");

                foreach (var curr in currencies)
                {
                    // Check if a rate exists for this exact date
                    string exactSql = "SELECT EntryID, Currency, ExchRate FROM CurrencyExchangeRates WHERE DT = @Date AND Currency = @Curr";
                    var exactRate = await db.QueryFirstOrDefaultAsync<CurrencyExchangeRateModel>(exactSql, new { Date = formattedDate, Curr = curr });
                    
                    if (exactRate != null)
                    {
                        resultList.Add(exactRate);
                    }
                    else
                    {
                        // Fallback to VCurrencyExchangeRates (most recent rate)
                        string fallbackSql = "SELECT TOP 1 ExchRate FROM VCurrencyExchangeRates WHERE Currency = @Curr";
                        float fallbackRate = await db.QueryFirstOrDefaultAsync<float>(fallbackSql, new { Curr = curr });
                        
                        resultList.Add(new CurrencyExchangeRateModel
                        {
                            EntryID = 0,
                            Currency = curr,
                            ExchRate = fallbackRate
                        });
                    }
                }

                return resultList;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting currency exchange rates.");
                throw;
            }
        }

        public async Task<int> SaveRateAsync(DateTime date, string currency, float rate, int entryId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                var formattedDate = date.ToString("yyyy-MM-dd");

                if (entryId == 0)
                {
                    // Insert
                    string sql = @"
                        INSERT INTO CurrencyExchangeRates(DT, Currency, ExchRate) 
                        VALUES(@Date, @Currency, @Rate);
                        SELECT CAST(SCOPE_IDENTITY() as int);";
                    
                    var newId = await db.QuerySingleAsync<int>(sql, new { Date = formattedDate, Currency = currency, Rate = rate });
                    return newId;
                }
                else
                {
                    // Update
                    string sql = "UPDATE CurrencyExchangeRates SET ExchRate = @Rate WHERE EntryID = @EntryId";
                    await db.ExecuteAsync(sql, new { Rate = rate, EntryId = entryId });
                    return entryId;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving currency exchange rate.");
                throw;
            }
        }
    }
}
