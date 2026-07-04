using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Company;
using Impulse.Services.Company;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Pages.Company
{
    public partial class PortsTree : ComponentBase
    {
        [Inject] private ILocationsService LocationsService { get; set; } = default!;
        [Inject] private IBlazorContextMenuService ContextMenuService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavManager { get; set; } = default!;

        private List<CountryNodeModel> treeData = new();
        private bool isLoading = true;
        private string searchText = "";

        // Modal State
        private bool showModal = false;
        private bool isEditing = false;
        private string modalTitle = "";
        private string modalInputText = "";
        
        // Tracking context for save
        private NodeType currentNodeType;
        private CountryNodeModel? selectedCountry;
        private CityNodeModel? selectedCity;
        private PortNodeModel? selectedPort;

        private enum NodeType { None, Country, City, Port }

        private IEnumerable<CountryNodeModel> FilteredTreeData
        {
            get
            {
                if (string.IsNullOrWhiteSpace(searchText))
                    return treeData;

                var q = searchText.Trim().ToLower();
                var result = new List<CountryNodeModel>();

                foreach (var country in treeData)
                {
                    bool countryMatch = country.CountryName.ToLower().Contains(q);
                    
                    var matchingCities = new List<CityNodeModel>();
                    foreach (var city in country.Cities)
                    {
                        bool cityMatch = city.City.ToLower().Contains(q);
                        
                        var matchingPorts = new List<PortNodeModel>();
                        foreach (var port in city.Ports)
                        {
                            if (countryMatch || cityMatch || port.Port.ToLower().Contains(q))
                            {
                                matchingPorts.Add(port);
                            }
                        }

                        if (countryMatch || cityMatch || matchingPorts.Any())
                        {
                            // If city matched because of children, we only want to show those matched children (unless country or city matched, then show all)
                            var cityClone = new CityNodeModel
                            {
                                CityID = city.CityID,
                                City = city.City,
                                CountryName = city.CountryName,
                                IsExpanded = true,
                                Ports = (countryMatch || cityMatch) ? city.Ports : matchingPorts
                            };
                            matchingCities.Add(cityClone);
                        }
                    }

                    if (countryMatch || matchingCities.Any())
                    {
                        var countryClone = new CountryNodeModel
                        {
                            CountryName = country.CountryName,
                            OriginalCountryName = country.OriginalCountryName,
                            IsExpanded = true,
                            Cities = countryMatch ? country.Cities : matchingCities
                        };
                        result.Add(countryClone);
                    }
                }
                return result;
            }
        }

        protected override async Task OnInitializedAsync()
        {
            await LoadDataAsync();
        }

        private async Task LoadDataAsync()
        {
            isLoading = true;
            try
            {
                // Preserve expansion state if possible
                var oldState = new HashSet<string>();
                foreach (var c in treeData)
                {
                    if (c.IsExpanded) oldState.Add("C_" + c.CountryName);
                    foreach (var city in c.Cities)
                    {
                        if (city.IsExpanded) oldState.Add("CI_" + city.CityID);
                    }
                }

                treeData = await LocationsService.GetFullTreeAsync();

                if (oldState.Count > 0)
                {
                    foreach (var c in treeData)
                    {
                        c.IsExpanded = oldState.Contains("C_" + c.CountryName);
                        foreach (var city in c.Cities)
                        {
                            city.IsExpanded = oldState.Contains("CI_" + city.CityID);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to load data: {ex.Message}");
            }
            finally
            {
                isLoading = false;
            }
        }

        private async Task RefreshTree()
        {
            await LoadDataAsync();
            NotificationService.Notify(NotificationSeverity.Info, "Refreshed", "Tree data refreshed.");
        }

        private void ToggleCountry(CountryNodeModel country)
        {
            country.IsExpanded = !country.IsExpanded;
        }

        private void ToggleCity(CityNodeModel city)
        {
            city.IsExpanded = !city.IsExpanded;
        }

        private void ExpandAll()
        {
            foreach (var c in treeData)
            {
                c.IsExpanded = true;
                foreach (var city in c.Cities)
                {
                    city.IsExpanded = true;
                }
            }
        }

        private void CollapseAll()
        {
            foreach (var c in treeData)
            {
                c.IsExpanded = false;
                foreach (var city in c.Cities)
                {
                    city.IsExpanded = false;
                }
            }
        }

        // Context Menus
        private async Task ShowCountryMenu(MouseEventArgs e, CountryNodeModel country)
        {
            await ContextMenuService.ShowMenu("countryContextMenu", (int)e.ClientX, (int)e.ClientY, country);
        }

        private async Task ShowCityMenu(MouseEventArgs e, CityNodeModel city)
        {
            await ContextMenuService.ShowMenu("cityContextMenu", (int)e.ClientX, (int)e.ClientY, city);
        }

        private async Task ShowPortMenu(MouseEventArgs e, PortNodeModel port)
        {
            await ContextMenuService.ShowMenu("portContextMenu", (int)e.ClientX, (int)e.ClientY, port);
        }

        // Action Handlers
        private void AddNewCountry()
        {
            currentNodeType = NodeType.Country;
            isEditing = false;
            modalTitle = "Add New Country";
            modalInputText = "";
            showModal = true;
        }

        private void EditCountry(CountryNodeModel country)
        {
            currentNodeType = NodeType.Country;
            isEditing = true;
            selectedCountry = country;
            modalTitle = "Rename Country";
            modalInputText = country.CountryName;
            showModal = true;
        }

        private async Task ConfirmDeleteCountry(CountryNodeModel country)
        {
            try
            {
                var success = await LocationsService.DeleteCountryAsync(country.OriginalCountryName);
                if (success)
                {
                    NotificationService.Notify(NotificationSeverity.Success, "Deleted", "Country deleted successfully.");
                    await LoadDataAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to delete country: {ex.Message}");
            }
        }

        private void AddNewCity(CountryNodeModel country)
        {
            currentNodeType = NodeType.City;
            isEditing = false;
            selectedCountry = country;
            modalTitle = $"Add New City to {country.CountryName}";
            modalInputText = "";
            showModal = true;
        }

        private void EditCity(CityNodeModel city)
        {
            currentNodeType = NodeType.City;
            isEditing = true;
            selectedCity = city;
            modalTitle = "Rename City";
            modalInputText = city.City;
            showModal = true;
        }

        private async Task ConfirmDeleteCity(CityNodeModel city)
        {
            try
            {
                var success = await LocationsService.DeleteCityAsync(city.CityID);
                if (success)
                {
                    NotificationService.Notify(NotificationSeverity.Success, "Deleted", "City deleted successfully.");
                    await LoadDataAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to delete city: {ex.Message}");
            }
        }

        private void AddNewPort(CityNodeModel city)
        {
            currentNodeType = NodeType.Port;
            isEditing = false;
            selectedCity = city;
            modalTitle = $"Add New Port to {city.City}";
            modalInputText = "";
            showModal = true;
        }

        private void EditPort(PortNodeModel port)
        {
            currentNodeType = NodeType.Port;
            isEditing = true;
            selectedPort = port;
            modalTitle = "Rename Port";
            modalInputText = port.Port;
            showModal = true;
        }

        private async Task ConfirmDeletePort(PortNodeModel port)
        {
            try
            {
                var success = await LocationsService.DeletePortAsync(port.PortID);
                if (success)
                {
                    NotificationService.Notify(NotificationSeverity.Success, "Deleted", "Port deleted successfully.");
                    await LoadDataAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to delete port: {ex.Message}");
            }
        }

        private void CloseModal()
        {
            showModal = false;
        }

        private async Task SaveRecord()
        {
            if (string.IsNullOrWhiteSpace(modalInputText))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation", "Name cannot be empty.");
                return;
            }

            try
            {
                bool success = false;
                switch (currentNodeType)
                {
                    case NodeType.Country:
                        if (isEditing && selectedCountry != null)
                            success = await LocationsService.UpdateCountryAsync(selectedCountry.OriginalCountryName, modalInputText);
                        else
                            success = await LocationsService.InsertCountryAsync(modalInputText);
                        break;
                    case NodeType.City:
                        if (isEditing && selectedCity != null)
                            success = await LocationsService.UpdateCityAsync(selectedCity.CityID, modalInputText);
                        else if (!isEditing && selectedCountry != null)
                        {
                            success = await LocationsService.InsertCityAsync(selectedCountry.CountryName, modalInputText);
                            selectedCountry.IsExpanded = true; // ensure it's visible
                        }
                        break;
                    case NodeType.Port:
                        if (isEditing && selectedPort != null)
                            success = await LocationsService.UpdatePortAsync(selectedPort.PortID, modalInputText);
                        else if (!isEditing && selectedCity != null)
                        {
                            success = await LocationsService.InsertPortAsync(selectedCity.CityID, modalInputText);
                            selectedCity.IsExpanded = true;
                        }
                        break;
                }

                if (success)
                {
                    NotificationService.Notify(NotificationSeverity.Success, "Saved", "Record saved successfully.");
                    showModal = false;
                    await LoadDataAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to save: {ex.Message}");
            }
        }

        private void NavigateToDashboard()
        {
            NavManager.NavigateTo("/company");
        }

        private string GetCountryFlagCssClass(string countryName)
        {
            if (string.IsNullOrWhiteSpace(countryName)) return "";
            
            var name = countryName.Trim().ToLower();
            return name switch
            {
                "pak" or "pakistan" => "fi-pk",
                "japan" => "fi-jp",
                "brazil" => "fi-br",
                "canada" => "fi-ca",
                "denmark" => "fi-dk",
                "italy" => "fi-it",
                "australia" => "fi-au",
                "spain" => "fi-es",
                "usa" or "united states" or "america" => "fi-us",
                "uk" or "united kingdom" or "england" => "fi-gb",
                "philippine" or "philippines" => "fi-ph",
                "china" => "fi-cn",
                "india" => "fi-in",
                "germany" => "fi-de",
                "france" => "fi-fr",
                "uae" or "united arab emirates" or "dubai" => "fi-ae",
                "bangladesh" => "fi-bd",
                "sri lanka" => "fi-lk",
                "saudi arabia" => "fi-sa",
                "turkey" => "fi-tr",
                "malaysia" => "fi-my",
                "indonesia" => "fi-id",
                "singapore" => "fi-sg",
                "thailand" => "fi-th",
                "vietnam" => "fi-vn",
                "south korea" or "korea" => "fi-kr",
                "russia" => "fi-ru",
                "south africa" => "fi-za",
                "egypt" => "fi-eg",
                "netherlands" or "holland" => "fi-nl",
                "belgium" => "fi-be",
                "switzerland" => "fi-ch",
                "sweden" => "fi-se",
                "norway" => "fi-no",
                "finland" => "fi-fi",
                "portugal" => "fi-pt",
                "mexico" => "fi-mx",
                "argentina" => "fi-ar",
                "chile" => "fi-cl",
                "colombia" => "fi-co",
                "peru" => "fi-pe",
                "new zealand" => "fi-nz",
                _ => ""
            };
        }
    }
}
