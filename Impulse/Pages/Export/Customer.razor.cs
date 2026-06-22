using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Export;
using Microsoft.AspNetCore.Components;
using Impulse.Services;

namespace Impulse.Pages.Export
{
    public partial class Customer : ComponentBase, IDisposable
    {
        // ------------------------------------------------------------------ //
        //  Injected services
        // ------------------------------------------------------------------ //

        [Inject] private IForeignCustomerService _customerService { get; set; } = null!;
        [Inject] private IAuditService _auditService { get; set; } = null!;
        [Inject] private IDBHelperService _dbHelper { get; set; } = null!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private NavigationManager Navigation { get; set; } = null!;
        [Inject] private IVoucherService _voucherService { get; set; } = null!;

        // ------------------------------------------------------------------ //
        //  Parameters
        // ------------------------------------------------------------------ //

        [Parameter] public string? CustId { get; set; }   // "CustCode+Country" when editing

        // ------------------------------------------------------------------ //
        //  State
        // ------------------------------------------------------------------ //

        private string strComputerName = string.Empty;
        private string strUserName = string.Empty;
        private bool isSaving = false;
        private string activeTab = "company";

        // Model
        private ForeignCustomerViewModel model = new();

        // Dropdown data lists
        private List<GenericDropDownModel> Countries = new();
        private List<GenericDropDownModel> Currencies = new();
        private List<GenericDropDownModel> PaymentTermsList = new();
        private List<GenericDropDownModel> TradeTermsList = new();
        private List<GenericDropDownModel> DTFormats = new();
        private List<GenericDropDownModel> FinishingQualities = new();
        private List<GenericDropDownModel> ShippingMethods = new();
        private List<GenericDropDownModel> CustomerSources = new();
        private List<GenericDropDownModel> EventRepetitions = new();
        private List<GenericDropDownModel> BankCountries = new();

        // Selected values for BlazoredTypeahead controls (header row)
        private GenericDropDownModel? SelectedCountry = null;
        private GenericDropDownModel? SelectedCurrency = null;

        // Selected values for BlazoredTypeahead controls (company tab)
        private GenericDropDownModel? SelectedTradeTerms = null;
        private GenericDropDownModel? SelectedPaymentTerms = null;
        private GenericDropDownModel? SelectedFinishingQuality = null;
        private GenericDropDownModel? SelectedCustomerSource = null;
        private GenericDropDownModel? SelectedShippingMethod = null;

        // Bank tab
        private ForeignCustomerBankViewModel currentBank = new();
        private int editingBankIndex = -1;
        private List<int> removedBankIds = new();
        private GenericDropDownModel? SelectedBankCountry = null;

        // Shipping Addresses tab
        private string newShippingAddress = string.Empty;
        private List<int> removedAddressIds = new();
        private ForeignCustomerAddressViewModel? editingAddress = null;

        // Events tab
        private string newEventDescription = string.Empty;
        private DateTime newEventDate = DateTime.Today;
        private GenericDropDownModel? SelectedEventRepeat = null;
        private List<int> removedEventIds = new();

        // Invoice To tab
        private string newInvoiceTo = string.Empty;
        private List<int> removedInvoiceToIds = new();
        private ForeignCustomerInvoiceToViewModel? editingInvoiceTo = null;

        // Timer for auto-hiding messages
        private Timer? _timer;

        public void Dispose()
        {
            _timer?.Dispose();
        }

        // ------------------------------------------------------------------ //
        //  Initialisation
        // ------------------------------------------------------------------ //

        protected override async Task OnInitializedAsync()
        {
            strComputerName = _auditService.GetClientIpAddress();
            strUserName = _auditService.GetCurrentUserName();

            await LoadDropdowns();

            if (!string.IsNullOrEmpty(CustId))
            {
                await LoadCustomer(CustId);
            }
        }

        private async Task LoadDropdowns()
        {
            Countries = await _dbHelper.GetListAsync<GenericDropDownModel>(
                "DISTINCT CountryName AS DropDownValue_Description, CountryName AS DropDownValue_ID",
                "Countries",
                "ORDER BY CountryName");

            Currencies = await _dbHelper.GetListAsync<GenericDropDownModel>(
                "DISTINCT Curr AS DropDownValue_Description, Curr AS DropDownValue_ID",
                "ForeignCustomers",
                string.Empty);

            PaymentTermsList = await _voucherService.GetValuesForDropDown(
                "PaymentTerms", "ID", "Description");

            TradeTermsList = await _voucherService.GetValuesForDropDown(
                "TradeTerms", "ID", "Description");

            FinishingQualities = await _voucherService.GetValuesForDropDown(
                "ItemsFinQuality", "ID", "Description", " ORDER BY Description");

            ShippingMethods = await _voucherService.GetValuesForDropDown(
                "ShippingMethods", "EntryID", "Description", " ORDER BY EntryID");

            CustomerSources = await _customerService.GetCustomerSourcesAsync();
            EventRepetitions = await _customerService.GetEventRepetitionsAsync();

            BankCountries = await _dbHelper.GetListAsync<GenericDropDownModel>(
                "CountryName AS DropDownValue_Description, CountryName AS DropDownValue_ID",
                "Countries",
                "WHERE CountryName<>'Pak' ORDER BY CountryName");

            DTFormats = new List<GenericDropDownModel>
            {
                new() { DropDownValue_ID = "dd-MM-yyyy",    DropDownValue_Description = "dd-MM-yyyy" },
                new() { DropDownValue_ID = "dd-MMM-yyyy",   DropDownValue_Description = "dd-MMM-yyyy" },
                new() { DropDownValue_ID = "MM-dd-yy",      DropDownValue_Description = "MM-dd-yy" },
                new() { DropDownValue_ID = "MMM-dd-yyyy",   DropDownValue_Description = "MMM-dd-yyyy" },
                new() { DropDownValue_ID = "MMMM-dd-yyyy",  DropDownValue_Description = "MMMM-dd-yyyy" },
                new() { DropDownValue_ID = "yy-MM-dd",      DropDownValue_Description = "yy-MM-dd" },
                new() { DropDownValue_ID = "yyyy-MM-dd",    DropDownValue_Description = "yyyy-MM-dd" },
                new() { DropDownValue_ID = "yyyy-MMM-dd",   DropDownValue_Description = "yyyy-MMM-dd" }
            };
        }

        private async Task LoadCustomer(string custId)
        {
            var loaded = await _customerService.GetCustomerAsync(custId);
            if (loaded == null) return;

            model = loaded;

            SelectedCountry = Countries.FirstOrDefault(c =>
                c.DropDownValue_ID == model.Country);
            SelectedCurrency = Currencies.FirstOrDefault(c =>
                c.DropDownValue_ID == model.Curr);

            // Initialize company tab selections
            SelectedTradeTerms = TradeTermsList.FirstOrDefault(t =>
                t.DropDownValue_Description == model.TradeTerms);
            SelectedPaymentTerms = PaymentTermsList.FirstOrDefault(p =>
                p.DropDownValue_Description == model.PaymentTerms);
            SelectedFinishingQuality = FinishingQualities.FirstOrDefault(f =>
                f.DropDownValue_Description == model.FinishingQuality);
            SelectedCustomerSource = CustomerSources.FirstOrDefault(s =>
                s.DropDownValue_Description == model.CustomerSource);
            SelectedShippingMethod = ShippingMethods.FirstOrDefault(sh =>
                sh.DropDownValue_Description == model.DefaultShipMethod);
        }

        // ------------------------------------------------------------------ //
        //  Tab switching
        // ------------------------------------------------------------------ //

        private void SetTab(string tab) => activeTab = tab;

        // ------------------------------------------------------------------ //
        //  Typeahead search helpers
        // ------------------------------------------------------------------ //

        private async Task<IEnumerable<GenericDropDownModel>> SearchCountries(string search)
            => await Task.FromResult(Countries
                .Where(x => x.DropDownValue_Description.Contains(search, StringComparison.OrdinalIgnoreCase))
                .ToList());

        private async Task<IEnumerable<GenericDropDownModel>> SearchCurrencies(string search)
            => await Task.FromResult(Currencies
                .Where(x => x.DropDownValue_Description.Contains(search, StringComparison.OrdinalIgnoreCase))
                .ToList());

        private async Task<IEnumerable<GenericDropDownModel>> SearchPaymentTerms(string search)
            => await Task.FromResult(PaymentTermsList
                .Where(x => x.DropDownValue_Description.Contains(search, StringComparison.OrdinalIgnoreCase))
                .ToList());

        private async Task<IEnumerable<GenericDropDownModel>> SearchTradeTerms(string search)
            => await Task.FromResult(TradeTermsList
                .Where(x => x.DropDownValue_Description.Contains(search, StringComparison.OrdinalIgnoreCase))
                .ToList());

        private async Task<IEnumerable<GenericDropDownModel>> SearchFinishingQualities(string search)
            => await Task.FromResult(FinishingQualities
                .Where(x => x.DropDownValue_Description.Contains(search, StringComparison.OrdinalIgnoreCase))
                .ToList());

        private async Task<IEnumerable<GenericDropDownModel>> SearchShippingMethods(string search)
            => await Task.FromResult(ShippingMethods
                .Where(x => x.DropDownValue_Description.Contains(search, StringComparison.OrdinalIgnoreCase))
                .ToList());

        private async Task<IEnumerable<GenericDropDownModel>> SearchCustomerSources(string search)
            => await Task.FromResult(CustomerSources
                .Where(x => x.DropDownValue_Description.Contains(search, StringComparison.OrdinalIgnoreCase))
                .ToList());

        private async Task<IEnumerable<GenericDropDownModel>> SearchBankCountries(string search)
            => await Task.FromResult(BankCountries
                .Where(x => x.DropDownValue_Description.Contains(search, StringComparison.OrdinalIgnoreCase))
                .ToList());

        private async Task<IEnumerable<GenericDropDownModel>> SearchEventRepetitions(string search)
            => await Task.FromResult(EventRepetitions
                .Where(x => x.DropDownValue_Description.Contains(search, StringComparison.OrdinalIgnoreCase))
                .ToList());

        // ------------------------------------------------------------------ //
        //  Selected-result callbacks
        // ------------------------------------------------------------------ //

        private void OnCountryChanged(GenericDropDownModel? val)
        {
            SelectedCountry = val;
            model.Country = val?.DropDownValue_ID ?? string.Empty;
        }

        private void OnCurrencyChanged(GenericDropDownModel? val)
        {
            SelectedCurrency = val;
            model.Curr = val?.DropDownValue_ID ?? string.Empty;
        }

        private void OnBankCountryChanged(GenericDropDownModel? val)
        {
            SelectedBankCountry = val;
            currentBank.Country = val?.DropDownValue_Description ?? string.Empty;
        }

        private void OnTradeTermsChanged(GenericDropDownModel? val)
        {
            SelectedTradeTerms = val;
            model.TradeTerms = val?.DropDownValue_Description ?? string.Empty;
        }

        private void OnPaymentTermsChanged(GenericDropDownModel? val)
        {
            SelectedPaymentTerms = val;
            model.PaymentTerms = val?.DropDownValue_Description ?? string.Empty;
        }

        private void OnFinishingQualityChanged(GenericDropDownModel? val)
        {
            SelectedFinishingQuality = val;
            model.FinishingQuality = val?.DropDownValue_Description ?? string.Empty;
        }

        private void OnCustomerSourceChanged(GenericDropDownModel? val)
        {
            SelectedCustomerSource = val;
            model.CustomerSource = val?.DropDownValue_Description ?? string.Empty;
        }

        private void OnShippingMethodChanged(GenericDropDownModel? val)
        {
            SelectedShippingMethod = val;
            model.DefaultShipMethod = val?.DropDownValue_Description ?? string.Empty;
        }

        // ------------------------------------------------------------------ //
        //  Bank Tab actions
        // ------------------------------------------------------------------ //

        private void AddOrUpdateBank()
        {
            if (string.IsNullOrWhiteSpace(currentBank.BankName))
            {
                NotificationServiceManager.ShowError("Validation", "Bank name is required.");
                return;
            }

            if (editingBankIndex >= 0)
            {
                model.Banks[editingBankIndex] = currentBank;
            }
            else
            {
                model.Banks.Add(currentBank);
            }

            ResetBankForm();
        }

        private void EditBank(int index)
        {
            editingBankIndex = index;
            currentBank = new ForeignCustomerBankViewModel
            {
                BankID = model.Banks[index].BankID,
                BankName = model.Banks[index].BankName,
                Country = model.Banks[index].Country,
                Phone1 = model.Banks[index].Phone1,
                Phone2 = model.Banks[index].Phone2,
                Fax1 = model.Banks[index].Fax1,
                Fax2 = model.Banks[index].Fax2,
                Email1 = model.Banks[index].Email1,
                Email2 = model.Banks[index].Email2,
                Address1 = model.Banks[index].Address1,
                Address2 = model.Banks[index].Address2,
                DefaultBank = model.Banks[index].DefaultBank
            };
            SelectedBankCountry = BankCountries.FirstOrDefault(c =>
                c.DropDownValue_Description == currentBank.Country);
        }

        private void RemoveBank(int index)
        {
            if (model.Banks[index].BankID > 0)
                removedBankIds.Add(model.Banks[index].BankID);
            model.Banks.RemoveAt(index);
            if (editingBankIndex == index) ResetBankForm();
        }

        private void ResetBankForm()
        {
            currentBank = new ForeignCustomerBankViewModel();
            SelectedBankCountry = null;
            editingBankIndex = -1;
        }

        private void SetDefaultBank(int index)
        {
            for (int i = 0; i < model.Banks.Count; i++)
                model.Banks[i].DefaultBank = (i == index);
        }

        // ------------------------------------------------------------------ //
        //  Shipping Addresses tab
        // ------------------------------------------------------------------ //

        private void AddShippingAddress()
        {
            if (string.IsNullOrWhiteSpace(newShippingAddress))
            {
                NotificationServiceManager.ShowError("Validation", "Please enter a shipping address.");
                return;
            }
            model.ShippingAddresses.Add(new ForeignCustomerAddressViewModel { ShippingAddress = newShippingAddress });
            newShippingAddress = string.Empty;
        }

        private void RemoveShippingAddress(ForeignCustomerAddressViewModel addr)
        {
            if (addr.EntryID > 0) removedAddressIds.Add(addr.EntryID);
            model.ShippingAddresses.Remove(addr);
        }

        private void StartEditAddress(ForeignCustomerAddressViewModel addr)
        {
            editingAddress = addr;
        }

        private void SaveEditAddress()
        {
            editingAddress = null;
        }

        // ------------------------------------------------------------------ //
        //  Events tab
        // ------------------------------------------------------------------ //

        private void AddEvent()
        {
            if (string.IsNullOrWhiteSpace(newEventDescription))
            {
                NotificationServiceManager.ShowError("Validation", "Please enter event description.");
                return;
            }
            if (SelectedEventRepeat == null)
            {
                NotificationServiceManager.ShowError("Validation", "Please select repetition.");
                return;
            }

            model.Events.Add(new ForeignCustomerEventViewModel
            {
                Description = newEventDescription,
                EventDT = newEventDate,
                RepeatedType = int.Parse(SelectedEventRepeat.DropDownValue_ID),
                RepeatDescription = SelectedEventRepeat.DropDownValue_Description
            });

            newEventDescription = string.Empty;
            newEventDate = DateTime.Today;
            SelectedEventRepeat = null;
        }

        private void RemoveEvent(ForeignCustomerEventViewModel ev)
        {
            if (ev.EntryID > 0) removedEventIds.Add(ev.EntryID);
            model.Events.Remove(ev);
        }

        // ------------------------------------------------------------------ //
        //  Invoice To tab
        // ------------------------------------------------------------------ //

        private void AddInvoiceTo()
        {
            if (string.IsNullOrWhiteSpace(newInvoiceTo))
            {
                NotificationServiceManager.ShowError("Validation", "Please enter Invoice To value.");
                return;
            }
            model.InvoiceTo.Add(new ForeignCustomerInvoiceToViewModel { InvoiceTo = newInvoiceTo });
            newInvoiceTo = string.Empty;
        }

        private void RemoveInvoiceTo(ForeignCustomerInvoiceToViewModel it)
        {
            if (it.EntryID > 0) removedInvoiceToIds.Add(it.EntryID);
            model.InvoiceTo.Remove(it);
        }

        private void StartEditInvoiceTo(ForeignCustomerInvoiceToViewModel it)
        {
            editingInvoiceTo = it;
        }

        private void SaveEditInvoiceTo()
        {
            editingInvoiceTo = null;
        }

        // ------------------------------------------------------------------ //
        //  Save & Cancel
        // ------------------------------------------------------------------ //

        private async Task SaveAndNew()
        {
            if (await SaveCustomer())
                ResetForm();
        }

        private async Task SaveAndClose()
        {
            if (await SaveCustomer())
                Navigation.NavigateTo("/export/customers");
        }

        private async Task<bool> SaveCustomer()
        {
            if (isSaving) return false;

            if (string.IsNullOrWhiteSpace(model.Name))
            {
                NotificationServiceManager.ShowError("Validation", "Customer name is required.");
                return false;
            }
            if (string.IsNullOrWhiteSpace(model.Country))
            {
                NotificationServiceManager.ShowError("Validation", "Country is required.");
                return false;
            }
            if (string.IsNullOrWhiteSpace(model.DTFormat))
            {
                NotificationServiceManager.ShowError("Validation", "Please select a date format.");
                return false;
            }

            try
            {
                isSaving = true;
                await _customerService.SaveCustomerAsync(
                    model, CustId ?? string.Empty,
                    strUserName, strComputerName,
                    removedBankIds, removedAddressIds, removedEventIds, removedInvoiceToIds);

                NotificationServiceManager.ShowSuccess("Saved", "Customer saved successfully.");
                return true;
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error", ex.Message);
                return false;
            }
            finally
            {
                isSaving = false;
            }
        }

        private void ResetForm()
        {
            model = new ForeignCustomerViewModel();
            SelectedCountry = null;
            SelectedCurrency = null;
            SelectedTradeTerms = null;
            SelectedPaymentTerms = null;
            SelectedFinishingQuality = null;
            SelectedCustomerSource = null;
            SelectedShippingMethod = null;
            currentBank = new ForeignCustomerBankViewModel();
            SelectedBankCountry = null;
            editingBankIndex = -1;
            removedBankIds.Clear();
            removedAddressIds.Clear();
            removedEventIds.Clear();
            removedInvoiceToIds.Clear();
            newShippingAddress = string.Empty;
            newEventDescription = string.Empty;
            newEventDate = DateTime.Today;
            SelectedEventRepeat = null;
            newInvoiceTo = string.Empty;
            editingAddress = null;
            editingInvoiceTo = null;
            activeTab = "company";
            CustId = null;
        }

        private void Cancel()
        {
            Navigation.NavigateTo("/export/customers");
        }
    }
}
