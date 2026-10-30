using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Common;
using DataAccessLibrary.Interface.QMS;
using DataAccessLibrary.Models.QMS;
using Impulse.Services.Setup;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.QMS.CustomerComplaints
{
    public partial class CustomerComplaintEntry : ComponentBase
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

        private List<CompanyLookupDto> companies = new();
        private List<string> customerList = new();
        private List<DepartmentLookupDto> departments = new();
        private List<QMSReasonLookupDto> qmsReasons = new();
        private List<EmployeeLookupDto> employeesByDept = new();
        private List<string> invoiceList = new();
        private List<InvoiceItemLookupDto> invoiceItems = new();
        private List<string> availableLots = new();

        // New Item Line Temp Inputs
        private string newDetailInvoice = "<All Items>";
        private InvoiceItemLookupDto? selectedInvoiceItem;
        private int newDetailQty = 1;
        private string newDetailReasonMain = string.Empty;
        private string newDetailLotRef = string.Empty;
        private string newDetailReason = string.Empty;

        protected override async Task OnInitializedAsync()
        {
            isLoading = true;
            try
            {
                await PermissionService.InitializeAsync();
                await LoadInitialLookupsAsync();

                if (IsEditMode)
                {
                    var existing = await ComplaintData.GetComplaintByIdAsync(Id);
                    if (existing != null)
                    {
                        model = existing;
                        var atts = await AttachmentData.GetAttachmentsAsync("CustomerComplaint", Id.ToString());
                        attachmentCount = atts.Count;
                        if (!string.IsNullOrWhiteSpace(model.Unit))
                        {
                            employeesByDept = await ComplaintData.GetEmployeesByDeptAsync(model.Unit);
                        }
                        if (!string.IsNullOrWhiteSpace(model.CustCode))
                        {
                            invoiceList = await ComplaintData.GetInvoicesByCustomerAsync(model.CustCode, model.Country);
                            invoiceItems = await ComplaintData.GetInvoiceItemsAsync(model.CustCode, model.Country, null);
                        }
                    }
                    else
                    {
                        NotificationServiceManager.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Error,
                            Summary = "Not Found",
                            Detail = $"Customer Complaint #{Id} could not be found.",
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
                        Company_RefID = companies.FirstOrDefault()?.EntryID ?? 1,
                        DT = DateTime.Today,
                        ComplaintNo = await ComplaintData.GetNextComplaintNoAsync(DateTime.Today)
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

        private async Task LoadInitialLookupsAsync()
        {
            companies = await ComplaintData.GetCompaniesAsync();
            customerList = await ComplaintData.GetCustomerCodesAsync();
            departments = await ComplaintData.GetDepartmentsAsync();
            qmsReasons = await ComplaintData.GetQMSReasonsAsync();
        }

        private void SetActiveTab(string tab)
        {
            activeTab = tab;
        }

        private void HandleAttachmentCountChanged(int count)
        {
            attachmentCount = count;
        }

        private async Task<IEnumerable<string>> SearchCustomers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return customerList;
            }

            return await Task.FromResult(
                customerList.Where(c => c.Contains(searchText, StringComparison.OrdinalIgnoreCase))
            );
        }

        private async Task OnCustomerChanged(string? customer)
        {
            model.CustCode = customer ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(model.CustCode))
            {
                var countries = await ComplaintData.GetCountriesAsync(model.CustCode);
                model.Country = countries.FirstOrDefault() ?? string.Empty;

                invoiceList = await ComplaintData.GetInvoicesByCustomerAsync(model.CustCode, model.Country);
                newDetailInvoice = "<All Items>";
                await ReloadInvoiceItemsAsync();
            }
            else
            {
                model.Country = string.Empty;
                invoiceList.Clear();
                invoiceItems.Clear();
            }
        }

        private async Task OnDepartmentChanged(ChangeEventArgs e)
        {
            model.Unit = e.Value?.ToString();
            if (!string.IsNullOrWhiteSpace(model.Unit))
            {
                employeesByDept = await ComplaintData.GetEmployeesByDeptAsync(model.Unit);
            }
            else
            {
                employeesByDept.Clear();
            }
            model.EmailTo = null;
            model.EmailCc = null;
        }

        private void OnPenaltyChanged(ChangeEventArgs e)
        {
            if (byte.TryParse(e.Value?.ToString(), out var penaltyVal))
            {
                model.Penalty = penaltyVal;
                if (model.Penalty != 1) model.PenaltyQty = null;
                if (model.Penalty != 2)
                {
                    model.PenaltyAmount = null;
                    model.PenaltyCurrency = null;
                }
            }
        }

        private async Task OnComplaintDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out var dt))
            {
                model.DT = dt;
                if (!IsEditMode)
                {
                    model.ComplaintNo = await ComplaintData.GetNextComplaintNoAsync(dt);
                }
            }
        }

        private async Task OnInvoiceSelectionChanged(ChangeEventArgs e)
        {
            newDetailInvoice = e.Value?.ToString() ?? "<All Items>";
            selectedInvoiceItem = null;
            availableLots.Clear();
            await ReloadInvoiceItemsAsync();
        }

        private async Task ReloadInvoiceItemsAsync()
        {
            if (string.IsNullOrWhiteSpace(model.CustCode))
            {
                invoiceItems.Clear();
                return;
            }

            var inv = newDetailInvoice == "<All Items>" ? null : newDetailInvoice;
            invoiceItems = await ComplaintData.GetInvoiceItemsAsync(model.CustCode, model.Country, inv);
        }

        private async Task<IEnumerable<InvoiceItemLookupDto>> SearchInvoiceItems(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return invoiceItems;
            }

            return await Task.FromResult(
                invoiceItems.Where(i => i.DisplayName.Contains(searchText, StringComparison.OrdinalIgnoreCase))
            );
        }

        private async Task OnInvoiceItemChanged(InvoiceItemLookupDto? item)
        {
            selectedInvoiceItem = item;
            availableLots.Clear();

            if (selectedInvoiceItem != null && !string.IsNullOrWhiteSpace(newDetailInvoice) && newDetailInvoice != "<All Items>")
            {
                availableLots = await ComplaintData.GetLotsAsync(newDetailInvoice, selectedInvoiceItem.ItemID);
            }
        }

        private void AddDetailLine()
        {
            if (selectedInvoiceItem == null)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Product Required",
                    Detail = "Please select a product / item for this defect line.",
                    Duration = 3000
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(newDetailReasonMain))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Reason Required",
                    Detail = "Please select a QMS Reason for this item.",
                    Duration = 3000
                });
                return;
            }

            if (newDetailQty <= 0)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Invalid Quantity",
                    Detail = "Quantity must be greater than zero.",
                    Duration = 3000
                });
                return;
            }

            var detail = new CustomerComplaintDetailModel
            {
                InvoiceNo = newDetailInvoice == "<All Items>" ? null : newDetailInvoice,
                ItemID = selectedInvoiceItem.CompItemID ?? selectedInvoiceItem.ItemID,
                ItemCode = selectedInvoiceItem.ItemID,
                Qty = newDetailQty,
                ItemReason_Main = newDetailReasonMain,
                EmailReference = newDetailLotRef,
                ItemReason = newDetailReason
            };

            model.Details.Add(detail);

            // Reset line inputs
            selectedInvoiceItem = null;
            newDetailQty = 1;
            newDetailReason = string.Empty;
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

        private async Task SaveComplaint(bool saveAndNew)
        {
            if (string.IsNullOrWhiteSpace(model.CustCode))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Customer Required",
                    Detail = "Please select a Customer before saving.",
                    Duration = 3500
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(model.ComplaintNo))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Complaint # Required",
                    Detail = "Please enter a Complaint Number.",
                    Duration = 3500
                });
                return;
            }

            if (!model.DT.HasValue)
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Date Required",
                    Detail = "Please choose a Complaint Date.",
                    Duration = 3500
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(model.Unit))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Department Required",
                    Detail = "Please select a Unit / Department.",
                    Duration = 3500
                });
                return;
            }

            if (!model.Details.Any(d => !d.IsDeleted))
            {
                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Items Required",
                    Detail = "Please add at least one defective item to the list.",
                    Duration = 3500
                });
                return;
            }

            isSaving = true;
            try
            {
                var currentUserName = PermissionService.CurrentUserName;
                int savedId = await ComplaintData.SaveComplaintAsync(model, model.Details, currentUserName, Environment.MachineName);

                NotificationServiceManager.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Saved Successfully",
                    Detail = $"Customer Complaint {model.ComplaintNo} saved successfully.",
                    Duration = 3000
                });

                if (saveAndNew)
                {
                    model = new CustomerComplaintModel
                    {
                        Company_RefID = companies.FirstOrDefault()?.EntryID ?? 1,
                        DT = DateTime.Today,
                        ComplaintNo = await ComplaintData.GetNextComplaintNoAsync(DateTime.Today)
                    };
                    selectedInvoiceItem = null;
                    newDetailInvoice = "<All Items>";
                    employeesByDept.Clear();
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

