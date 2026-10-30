using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Common;
using DataAccessLibrary.Interface.QMS;
using DataAccessLibrary.Models.QMS;
using Impulse.Services.Setup;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;

namespace Impulse.Pages.QMS.CustomerFeedback
{
    public partial class CustomerFeedbackEntry : ComponentBase
    {
        [Parameter] public int Id { get; set; }
        [SupplyParameterFromQuery(Name = "returnUrl")] public string? ReturnUrl { get; set; }

        [Inject] private ICustomerComplaintDataAccess ComplaintData { get; set; } = null!;
        [Inject] private IEntityAttachmentDataAccess AttachmentData { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private NotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IUserPermissionService PermissionService { get; set; } = null!;

        private bool IsEditMode => Id > 0;
        private string activeTab = "details";
        private bool isLoading = true;
        private bool isSaving = false;
        private int attachmentCount = 0;

        private CustomerComplaintModel model = new();

        private List<string> customerList = new();
        private List<string> countryList = new();
        private List<string> invoiceList = new();
        private List<InvoiceItemLookupDto> invoiceItems = new();

        // New Item Line Temp Inputs
        private string newDetailInvoice = "<All Items>";
        private InvoiceItemLookupDto? selectedInvoiceItem;
        private int newDetailQty = 1;
        private string newDetailLotRef = string.Empty;

        protected override async Task OnInitializedAsync()
        {
            isLoading = true;
            try
            {
                await PermissionService.InitializeAsync();
                customerList = await ComplaintData.GetCustomerCodesAsync();

                if (IsEditMode)
                {
                    var existing = await ComplaintData.GetComplaintByIdAsync(Id);
                    if (existing != null)
                    {
                        model = existing;
                        if (!string.IsNullOrWhiteSpace(model.CustCode))
                        {
                            countryList = await ComplaintData.GetCountriesAsync(model.CustCode);
                            invoiceList = await ComplaintData.GetInvoicesByCustomerAsync(model.CustCode, model.Country);
                            invoiceItems = await ComplaintData.GetInvoiceItemsAsync(model.CustCode, model.Country, null);
                        }

                        var atts = await AttachmentData.GetAttachmentsAsync("CustomerComplaint", Id.ToString());
                        attachmentCount = atts.Count;
                    }
                    else
                    {
                        NotificationServiceManager.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Error,
                            Summary = "Not Found",
                            Detail = $"Customer Feedback #{Id} could not be found.",
                            Duration = 4000
                        });
                        Cancel();
                        return;
                    }
                }
                else
                {
                    model = new CustomerComplaintModel
                    {
                        Type = 1,
                        DT = DateTime.Today,
                        ComplaintNo = await ComplaintData.GetNextFeedbackNoAsync(DateTime.Today)
                    };
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Form",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                isLoading = false;
            }
        }

        private void SetActiveTab(string tab)
        {
            activeTab = tab;
        }

        private async Task<IEnumerable<string>> SearchCustomers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return await Task.FromResult(customerList);
            }

            return await Task.FromResult(
                customerList.Where(c => c.Contains(searchText, StringComparison.OrdinalIgnoreCase))
            );
        }

        private async Task OnCustomerChanged(string? selectedCust)
        {
            model.CustCode = selectedCust ?? string.Empty;
            selectedInvoiceItem = null;
            newDetailInvoice = "<All Items>";

            if (!string.IsNullOrWhiteSpace(model.CustCode))
            {
                countryList = await ComplaintData.GetCountriesAsync(model.CustCode);
                model.Country = countryList.FirstOrDefault() ?? string.Empty;

                invoiceList = await ComplaintData.GetInvoicesByCustomerAsync(model.CustCode, model.Country);
                invoiceItems = await ComplaintData.GetInvoiceItemsAsync(model.CustCode, model.Country, null);
            }
            else
            {
                countryList.Clear();
                model.Country = string.Empty;
                invoiceList.Clear();
                invoiceItems.Clear();
            }
        }

        private async Task OnCountryChanged(ChangeEventArgs e)
        {
            model.Country = e.Value?.ToString() ?? string.Empty;
            selectedInvoiceItem = null;
            newDetailInvoice = "<All Items>";

            if (!string.IsNullOrWhiteSpace(model.CustCode))
            {
                invoiceList = await ComplaintData.GetInvoicesByCustomerAsync(model.CustCode, model.Country);
                invoiceItems = await ComplaintData.GetInvoiceItemsAsync(model.CustCode, model.Country, null);
            }
        }

        private async Task OnFeedbackDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out var dt))
            {
                model.DT = dt;
                if (!IsEditMode)
                {
                    model.ComplaintNo = await ComplaintData.GetNextFeedbackNoAsync(dt);
                }
            }
        }

        private async Task OnInvoiceChanged(ChangeEventArgs e)
        {
            newDetailInvoice = e.Value?.ToString() ?? "<All Items>";
            selectedInvoiceItem = null;

            var invFilter = newDetailInvoice == "<All Items>" ? null : newDetailInvoice;
            invoiceItems = await ComplaintData.GetInvoiceItemsAsync(model.CustCode, model.Country, invFilter);
        }

        private async Task<IEnumerable<InvoiceItemLookupDto>> SearchInvoiceItems(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return await Task.FromResult(invoiceItems);
            }

            return await Task.FromResult(
                invoiceItems.Where(i =>
                    i.DisplayName.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                    (!string.IsNullOrEmpty(i.ItemCode) && i.ItemCode.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                )
            );
        }

        private void OnInvoiceItemChanged(InvoiceItemLookupDto? itm)
        {
            selectedInvoiceItem = itm;
        }

        private void HandleRefKeyDown(KeyboardEventArgs e)
        {
            if (e.Key == "Enter")
            {
                AddDetailLine();
            }
        }

        private void AddDetailLine()
        {
            if (selectedInvoiceItem == null || string.IsNullOrWhiteSpace(selectedInvoiceItem.ItemID))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Item Required",
                    Detail = "Please select a product/item before adding to list.",
                    Duration = 3500
                });
                return;
            }

            if (newDetailQty <= 0)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Invalid Quantity",
                    Detail = "Quantity must be at least 1.",
                    Duration = 3500
                });
                return;
            }

            model.Details.Add(new CustomerComplaintDetailModel
            {
                RefID = model.EntryID,
                InvoiceNo = newDetailInvoice == "<All Items>" ? null : newDetailInvoice,
                ItemID = selectedInvoiceItem.ItemID,
                ItemCode = selectedInvoiceItem.ItemCode,
                Qty = newDetailQty,
                EmailReference = string.IsNullOrWhiteSpace(newDetailLotRef) ? null : newDetailLotRef.Trim()
            });

            // Reset temp line inputs
            selectedInvoiceItem = null;
            newDetailQty = 1;
            newDetailLotRef = string.Empty;
        }

        private void RemoveDetailLine(CustomerComplaintDetailModel detail)
        {
            if (detail.EntryID > 0)
            {
                detail.IsDeleted = true;
            }
            else
            {
                model.Details.Remove(detail);
            }
        }

        private void HandleAttachmentCountChanged(int count)
        {
            attachmentCount = count;
        }

        private async Task SaveFeedback(bool saveAndNew)
        {
            if (string.IsNullOrWhiteSpace(model.CustCode))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Customer is required.",
                    Duration = 3500
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(model.ComplaintNo))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Feedback # is required.",
                    Duration = 3500
                });
                return;
            }

            isSaving = true;
            try
            {
                model.Type = 1; // 1: Feedback / Observation
                var userName = !string.IsNullOrWhiteSpace(PermissionService.CurrentUserName) ? PermissionService.CurrentUserName : "System";
                var machineName = Environment.MachineName;

                var savedId = await ComplaintData.SaveComplaintAsync(model, model.Details, userName, machineName);

                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Feedback Saved",
                    Detail = $"Customer Feedback {model.ComplaintNo} saved successfully.",
                    Duration = 3500
                });

                if (saveAndNew)
                {
                    model = new CustomerComplaintModel
                    {
                        Type = 1,
                        DT = DateTime.Today,
                        ComplaintNo = await ComplaintData.GetNextFeedbackNoAsync(DateTime.Today)
                    };
                    selectedInvoiceItem = null;
                    newDetailInvoice = "<All Items>";
                    activeTab = "details";
                }
                else
                {
                    NavigationManager.NavigateTo(ReturnUrl ?? "/qms/command-center");
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Save Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                isSaving = false;
            }
        }

        private void Cancel()
        {
            NavigationManager.NavigateTo(ReturnUrl ?? "/qms/command-center");
        }
    }
}

