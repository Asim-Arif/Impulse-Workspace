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
    public class LocationsDataAccess : ILocationsDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<LocationsDataAccess> _logger;

        public LocationsDataAccess(IConfiguration configuration, ILogger<LocationsDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
            _logger = logger;
        }

        public async Task<List<CountryNodeModel>> GetFullTreeAsync()
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                
                // Fetch all data
                var countries = await db.QueryAsync<CountryNodeModel>(
                    "SELECT CountryName, CountryName as OriginalCountryName FROM Countries ORDER BY CountryName");
                    
                var cities = await db.QueryAsync<CityNodeModel>(
                    "SELECT CityID, CountryName, City FROM Cities ORDER BY City");
                    
                var ports = await db.QueryAsync<PortNodeModel>(
                    "SELECT PortID, CityID, Port FROM Ports ORDER BY Port");

                // Build the tree in memory
                var countryList = countries.ToList();
                var cityList = cities.ToList();
                var portList = ports.ToList();

                foreach (var city in cityList)
                {
                    city.Ports = portList.Where(p => p.CityID == city.CityID).ToList();
                }

                foreach (var country in countryList)
                {
                    country.Cities = cityList.Where(c => c.CountryName == country.CountryName).ToList();
                }

                return countryList;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting locations tree.");
                throw;
            }
        }

        public async Task<bool> InsertCountryAsync(string countryName)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "INSERT INTO Countries (CountryName) VALUES (@CountryName)";
                var rowsAffected = await db.ExecuteAsync(sql, new { CountryName = countryName });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error inserting country.");
                throw;
            }
        }

        public async Task<bool> UpdateCountryAsync(string oldCountryName, string newCountryName)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "UPDATE Countries SET CountryName = @NewCountryName WHERE CountryName = @OldCountryName";
                var rowsAffected = await db.ExecuteAsync(sql, new { NewCountryName = newCountryName, OldCountryName = oldCountryName });
                
                // Note: If cascading updates are not enabled in SQL Server for the Cities FK, 
                // we would need to manually update Cities.CountryName here. 
                // Assuming standard legacy behavior.
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating country.");
                throw;
            }
        }

        public async Task<bool> DeleteCountryAsync(string countryName)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "DELETE FROM Countries WHERE CountryName = @CountryName";
                var rowsAffected = await db.ExecuteAsync(sql, new { CountryName = countryName });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting country.");
                throw;
            }
        }

        public async Task<bool> InsertCityAsync(string countryName, string cityName)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "INSERT INTO Cities (CountryName, City) VALUES (@CountryName, @City)";
                var rowsAffected = await db.ExecuteAsync(sql, new { CountryName = countryName, City = cityName });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error inserting city.");
                throw;
            }
        }

        public async Task<bool> UpdateCityAsync(int cityId, string cityName)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "UPDATE Cities SET City = @CityName WHERE CityID = @CityId";
                var rowsAffected = await db.ExecuteAsync(sql, new { CityName = cityName, CityId = cityId });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating city.");
                throw;
            }
        }

        public async Task<bool> DeleteCityAsync(int cityId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "DELETE FROM Cities WHERE CityID = @CityId";
                var rowsAffected = await db.ExecuteAsync(sql, new { CityId = cityId });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting city.");
                throw;
            }
        }

        public async Task<bool> InsertPortAsync(int cityId, string portName)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "INSERT INTO Ports (CityID, Port) VALUES (@CityId, @PortName)";
                var rowsAffected = await db.ExecuteAsync(sql, new { CityId = cityId, PortName = portName });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error inserting port.");
                throw;
            }
        }

        public async Task<bool> UpdatePortAsync(int portId, string portName)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "UPDATE Ports SET Port = @PortName WHERE PortID = @PortId";
                var rowsAffected = await db.ExecuteAsync(sql, new { PortName = portName, PortId = portId });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating port.");
                throw;
            }
        }

        public async Task<bool> DeletePortAsync(int portId)
        {
            try
            {
                using IDbConnection db = new SqlConnection(_connectionString);
                string sql = "DELETE FROM Ports WHERE PortID = @PortId";
                var rowsAffected = await db.ExecuteAsync(sql, new { PortId = portId });
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting port.");
                throw;
            }
        }
    }
}
