using DataAccessLibrary.Interface.Company;
using DataAccessLibrary.Models.ViewModels.Company;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Company
{
    public class LocationsService : ILocationsService
    {
        private readonly ILocationsDataAccess _dataAccess;

        public LocationsService(ILocationsDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public Task<List<CountryNodeModel>> GetFullTreeAsync()
        {
            return _dataAccess.GetFullTreeAsync();
        }

        public Task<bool> InsertCountryAsync(string countryName)
        {
            return _dataAccess.InsertCountryAsync(countryName);
        }

        public Task<bool> UpdateCountryAsync(string oldCountryName, string newCountryName)
        {
            return _dataAccess.UpdateCountryAsync(oldCountryName, newCountryName);
        }

        public Task<bool> DeleteCountryAsync(string countryName)
        {
            return _dataAccess.DeleteCountryAsync(countryName);
        }

        public Task<bool> InsertCityAsync(string countryName, string cityName)
        {
            return _dataAccess.InsertCityAsync(countryName, cityName);
        }

        public Task<bool> UpdateCityAsync(int cityId, string cityName)
        {
            return _dataAccess.UpdateCityAsync(cityId, cityName);
        }

        public Task<bool> DeleteCityAsync(int cityId)
        {
            return _dataAccess.DeleteCityAsync(cityId);
        }

        public Task<bool> InsertPortAsync(int cityId, string portName)
        {
            return _dataAccess.InsertPortAsync(cityId, portName);
        }

        public Task<bool> UpdatePortAsync(int portId, string portName)
        {
            return _dataAccess.UpdatePortAsync(portId, portName);
        }

        public Task<bool> DeletePortAsync(int portId)
        {
            return _dataAccess.DeletePortAsync(portId);
        }
    }
}
