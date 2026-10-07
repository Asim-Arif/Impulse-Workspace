using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.IntraOffice;
using Impulse.Services.IntraOffice;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;

namespace Impulse.Pages.IntraOffice.CRM
{
    public partial class Customer360 : ComponentBase
    {
        [Inject] private IIntraOfficeService IntraOfficeService { get; set; } = null!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = null!;
        [Inject] private NavigationManager Nav { get; set; } = null!;
        [Inject] private Radzen.NotificationService RadzenNotify { get; set; } = null!;

        [Parameter] public string? CustCode { get; set; }

        private bool isLoading = true;
        private string activeTab = "overview";
        private Customer360Dto? customer;
        private List<Customer360LookupDto> customerLookupList = new();

        private bool showAddContactModal = false;
        private CustomerContactDto? newContact;

        private bool showAddActivityModal = false;
        private CustomerActivityDto? newActivity;

        private bool showConvertModal = false;
        private ConvertCustomerToForeignCustomerModel? convertModel;

        private string currentUserName = "System";

        protected override async Task OnInitializedAsync()
        {
            var authState = await AuthStateProvider.GetAuthenticationStateAsync();
            currentUserName = authState.User.Identity?.Name ?? "Zeeshan";

            await LoadLookupList();
            await LoadCustomerData(CustCode);
        }

        protected override async Task OnParametersSetAsync()
        {
            if (customer != null && !string.Equals(customer.CustomerCode, CustCode, StringComparison.OrdinalIgnoreCase))
            {
                await LoadCustomerData(CustCode);
            }
        }

        private async Task LoadLookupList()
        {
            try
            {
                customerLookupList = await IntraOfficeService.GetCustomer360LookupListAsync();
            }
            catch (Exception ex)
            {
                RadzenNotify.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Lookup Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        private async Task LoadCustomerData(string? code)
        {
            isLoading = true;
            try
            {
                customer = await IntraOfficeService.GetCustomer360Async(code);
                if (customer != null)
                {
                    newContact = new CustomerContactDto { CustomerId = customer.Id };
                    newActivity = new CustomerActivityDto
                    {
                        CustomerId = customer.Id,
                        ActivityType = "Call",
                        PerformedBy = currentUserName
                    };
                }
            }
            catch (Exception ex)
            {
                RadzenNotify.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error Loading Customer",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                isLoading = false;
            }
        }

        private async Task LoadDefaultCustomer()
        {
            CustCode = "CUST-ME-2004";
            await LoadCustomerData("CUST-ME-2004");
        }

        private void OnCustomerSelected(ChangeEventArgs e)
        {
            var selectedCode = e.Value?.ToString();
            if (!string.IsNullOrWhiteSpace(selectedCode))
            {
                Nav.NavigateTo($"/office/customer-360/{selectedCode}");
            }
        }

        private async Task SaveContact()
        {
            if (customer == null || newContact == null || string.IsNullOrWhiteSpace(newContact.Name))
            {
                RadzenNotify.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Contact name is required.",
                    Duration = 3000
                });
                return;
            }

            customer.Contacts.Add(new CustomerContactDto
            {
                Id = customer.Contacts.Count + 1,
                CustomerId = customer.Id,
                Name = newContact.Name.Trim(),
                Designation = newContact.Designation,
                Email = newContact.Email,
                Phone = newContact.Phone,
                WhatsApp = newContact.Phone,
                IsPrimary = newContact.IsPrimary
            });

            await IntraOfficeService.AddCustomer360ContactAsync(newContact);

            showAddContactModal = false;
            newContact = new CustomerContactDto { CustomerId = customer.Id };

            RadzenNotify.Notify(new Radzen.NotificationMessage
            {
                Severity = Radzen.NotificationSeverity.Success,
                Summary = "Contact Saved",
                Detail = "Contact person was added successfully.",
                Duration = 3000
            });
        }

        private async Task SaveActivity()
        {
            if (customer == null || newActivity == null || string.IsNullOrWhiteSpace(newActivity.Subject))
            {
                RadzenNotify.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Activity subject is required.",
                    Duration = 3000
                });
                return;
            }

            var activityToSave = new CustomerActivityDto
            {
                Id = customer.Activities.Count + 1,
                CustomerId = customer.Id,
                ActivityType = newActivity.ActivityType,
                Subject = newActivity.Subject.Trim(),
                Description = newActivity.Description?.Trim() ?? "",
                PerformedBy = currentUserName,
                ActivityDate = DateTime.UtcNow
            };

            customer.Activities.Insert(0, activityToSave);
            await IntraOfficeService.AddCustomer360ActivityAsync(activityToSave);

            showAddActivityModal = false;
            newActivity = new CustomerActivityDto
            {
                CustomerId = customer.Id,
                ActivityType = "Call",
                PerformedBy = currentUserName
            };

            RadzenNotify.Notify(new Radzen.NotificationMessage
            {
                Severity = Radzen.NotificationSeverity.Success,
                Summary = "Activity Logged",
                Detail = "Customer communication activity logged.",
                Duration = 3000
            });
        }

        private void OpenConvertModal()
        {
            if (customer == null) return;

            var primaryContact = customer.Contacts.Find(c => c.IsPrimary) ?? (customer.Contacts.Count > 0 ? customer.Contacts[0] : null);

            convertModel = new ConvertCustomerToForeignCustomerModel
            {
                LeadId = customer.LeadId,
                CustCode = customer.CustomerCode,
                Name = customer.CompanyName,
                Country = customer.Country,
                City = customer.City,
                Address = customer.BillingAddress,
                Phone1 = primaryContact?.Phone,
                Email1 = primaryContact?.Email,
                URL = customer.Website,
                Curr = customer.PreferredCurrency,
                TradeTerms = customer.Incoterm,
                PaymentTerms = customer.PaymentTerms,
                AccNo = customer.TaxVatEoriNumber,
                Cont1name = primaryContact?.Name,
                Cont1Designation = primaryContact?.Designation,
                Cont1Email = primaryContact?.Email,
                cont1Phone = primaryContact?.Phone,
                Cont1Mobile = primaryContact?.WhatsApp,
                SpecialInstructions = customer.Notes,
                Customer_Type = customer.CustomerType,
                ConvertedBy = currentUserName
            };

            showConvertModal = true;
        }

        private async Task ExecuteConvertCustomer()
        {
            if (convertModel == null || string.IsNullOrWhiteSpace(convertModel.CustCode) || string.IsNullOrWhiteSpace(convertModel.Name))
            {
                RadzenNotify.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Validation Error",
                    Detail = "Customer code and company name are required.",
                    Duration = 3000
                });
                return;
            }

            try
            {
                var success = await IntraOfficeService.ConvertCustomer360ToForeignCustomerAsync(convertModel);
                if (success)
                {
                    if (customer != null)
                    {
                        customer.IsConvertedToForeignCustomer = true;
                        customer.ForeignCustomerCode = convertModel.CustCode;
                    }

                    showConvertModal = false;
                    await LoadLookupList();

                    RadzenNotify.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Success,
                        Summary = "Graduation Complete",
                        Detail = $"Customer '{convertModel.Name}' [{convertModel.CustCode}] is now registered as an Impulse Foreign Customer!",
                        Duration = 5000
                    });
                }
            }
            catch (Exception ex)
            {
                RadzenNotify.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Conversion Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }
    }
}
