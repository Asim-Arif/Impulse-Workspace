using DataAccessLibrary.Models.ViewModels.Company;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.Company
{
    public interface ILocationsService
    {
        Task<List<CountryNodeModel>> GetFullTreeAsync();
        
        Task<bool> InsertCountryAsync(string countryName);
        Task<bool> UpdateCountryAsync(string oldCountryName, string newCountryName);
        Task<bool> DeleteCountryAsync(string countryName);

        Task<bool> InsertCityAsync(string countryName, string cityName);
        Task<bool> UpdateCityAsync(int cityId, string cityName);
        Task<bool> DeleteCityAsync(int cityId);

        Task<bool> InsertPortAsync(int cityId, string portName);
        Task<bool> UpdatePortAsync(int portId, string portName);
        Task<bool> DeletePortAsync(int portId);
    }
}
