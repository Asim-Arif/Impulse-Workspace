using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Services.Export;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;

namespace Impulse.Components.Export.NewProformaInvoiceTabs
{
    public partial class OtherInfoTab
    {
        [Parameter] public ProformaInvoiceViewModel Proforma { get; set; } = default!;
        [Parameter] public bool InsuranceAgentDisabled { get; set; }
        [Parameter] public List<AgentLookupModel> ShippingAgents { get; set; } = new();
        [Parameter] public List<AgentLookupModel> InsuranceAgents { get; set; } = new();
        [Parameter] public List<CustomerBankModel> CustomerBanks { get; set; } = new();
        [Parameter] public List<CompanyBankAccountModel> CompanyBankAccounts { get; set; } = new();
        [Parameter] public List<ShippingAddressPresetModel> ShippingAddressPresets { get; set; } = new();
        [Parameter] public List<InvoiceToPresetModel> InvoiceToPresets { get; set; } = new();
        [Parameter] public EventCallback<string> OnCalcSteelWeight { get; set; }

        private string _selectedOurBankName = string.Empty;
        
        public string SelectedOurBankName 
        {
            get => _selectedOurBankName;
            set 
            {
                if (_selectedOurBankName != value)
                {
                    _selectedOurBankName = value;
                    var accounts = CompanyBankAccounts.Where(a => a.BankName == _selectedOurBankName).ToList();
                    if (accounts.Any())
                    {
                        Proforma.OurBankID = accounts.First().BankID;
                    }
                    else
                    {
                        Proforma.OurBankID = 0;
                    }
                }
            }
        }

        protected override void OnParametersSet()
        {
            base.OnParametersSet();
            if (Proforma.OurBankID > 0 && CompanyBankAccounts.Any())
            {
                var match = CompanyBankAccounts.FirstOrDefault(a => a.BankID == Proforma.OurBankID);
                if (match != null)
                {
                    _selectedOurBankName = match.BankName;
                }
            }
        }

        [Inject] private IProformaInvoiceService ProformaService { get; set; } = default!;

        // Port of Loading
        private List<string> LoadCountries { get; set; } = new() { "Pakistan" };
        private string LoadCountry { get; set; } = "Pakistan";
        private List<CityModel> LoadCities { get; set; } = new();
        private List<PortModel> LoadPorts { get; set; } = new();
        private int SelectedLoadCityId { get; set; }

        // Port of Discharge
        private List<string> DisCountries { get; set; } = new();
        private string DisCountry { get; set; } = string.Empty;
        private List<CityModel> DisCities { get; set; } = new();
        private List<PortModel> DisPorts { get; set; } = new();
        private int SelectedDisCityId { get; set; }

        // Bank address display
        private string BankAddress { get; set; } = string.Empty;

        protected override async Task OnInitializedAsync()
        {
            // Load cities for Pakistan (default load country)
            LoadCities = await ProformaService.GetCitiesAsync("Pakistan");
        }

        // Called by parent after customer/country selection to set the discharge country list
        public void SetDisCountryList(List<string> countries)
        {
            DisCountries = countries;
            if (countries.Count > 0 && string.IsNullOrEmpty(DisCountry))
                DisCountry = countries[0];
        }

        public async Task RestorePortSelections(int? loadPortId, int? disPortId)
        {
            if (loadPortId.HasValue && loadPortId > 0)
            {
                var loadDetails = await ProformaService.GetPortCityAndCountryAsync(loadPortId.Value);
                if (loadDetails.CityID > 0)
                {
                    LoadCountry = loadDetails.CountryName;
                    LoadCities = await ProformaService.GetCitiesAsync(LoadCountry);
                    SelectedLoadCityId = loadDetails.CityID;
                    LoadPorts = await ProformaService.GetPortsAsync(SelectedLoadCityId);
                }
                Proforma.LoadPort = loadPortId;
            }
            if (disPortId.HasValue && disPortId > 0)
            {
                var disDetails = await ProformaService.GetPortCityAndCountryAsync(disPortId.Value);
                if (disDetails.CityID > 0)
                {
                    DisCountry = disDetails.CountryName;
                    if (!DisCountries.Contains(DisCountry))
                    {
                        DisCountries.Add(DisCountry);
                    }
                    DisCities = await ProformaService.GetCitiesAsync(DisCountry);
                    SelectedDisCityId = disDetails.CityID;
                    DisPorts = await ProformaService.GetPortsAsync(SelectedDisCityId);
                }
                Proforma.DischargePort = disPortId;
            }
        }

        private async Task OnLoadCountryChanged()
        {
            LoadCities = await ProformaService.GetCitiesAsync(LoadCountry);
            LoadPorts.Clear();
            SelectedLoadCityId = 0;
            Proforma.LoadPort = null;
        }

        private async Task OnLoadCityChanged()
        {
            if (SelectedLoadCityId > 0)
                LoadPorts = await ProformaService.GetPortsAsync(SelectedLoadCityId);
            else
                LoadPorts.Clear();
            Proforma.LoadPort = null;
        }

        private async Task OnDisCountryInput(ChangeEventArgs e)
        {
            DisCountry = e.Value?.ToString() ?? string.Empty;
            DisCities = await ProformaService.GetCitiesAsync(DisCountry);
            DisPorts.Clear();
            SelectedDisCityId = 0;
            Proforma.DischargePort = null;
        }

        private async Task OnDisCityChanged()
        {
            if (SelectedDisCityId > 0)
                DisPorts = await ProformaService.GetPortsAsync(SelectedDisCityId);
            else
                DisPorts.Clear();
            Proforma.DischargePort = null;
        }

        private void OnCustomerBankChanged()
        {
            var bank = CustomerBanks.Find(b => b.BankID == Proforma.BankID);
            BankAddress = bank?.Address1 ?? string.Empty;
        }

        private void OnShippingAddressPresetChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int id))
            {
                var preset = ShippingAddressPresets.Find(sa => sa.EntryID == id);
                if (preset != null) Proforma.ShippingAddress = preset.ShippingAddress;
            }
        }

        private void OnInvoiceToPresetChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int id))
            {
                var preset = InvoiceToPresets.Find(it => it.EntryID == id);
                if (preset != null) Proforma.InvoiceTo = preset.InvoiceTo;
            }
        }

        private async Task CalcSteelWeight()
        {
            await OnCalcSteelWeight.InvokeAsync(Proforma.PInvoice);
        }
    }
}
