using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Services;
using Impulse.Services.Export;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.JSInterop;

namespace Impulse.Pages.Export.Invoices
{
    public partial class InvoiceList : ComponentBase
    {
        [Parameter] public string InvoiceTypeString { get; set; } = "custom";

        [Inject] private IInvoiceListService InvoiceListService { get; set; } = null!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IDBHelperService DbHelper { get; set; } = null!;
        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; } = null!;
        [Inject] private IJSRuntime JS { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;

        private int invoiceType = 0; // 0 = Custom, 1 = Commercial, 2 = Bank

        private bool isLoading = true;
        private string searchText = string.Empty;
        private string selectedDateRange = "180";
        private DateTime? fromDate = DateTime.Today.AddDays(-180);
        private DateTime? toDate = DateTime.Today;

        private long selectedCompanyId = 0;
        private string selectedCountry = string.Empty;
        private CustomerLookupModel? selectedCustomer;

        private List<CompanyLookupModel> companies = new();
        private List<CustomerLookupModel> customers = new();
        private List<string> countries = new();
        private List<InvoiceListModel> allItems = new();
        private InvoiceListModel? contextMenuRowItem;

        // Modals
        private bool showPasswordModal = false;
        private string passwordSettingName = string.Empty;
        private string enteredPassword = string.Empty;
        private string passwordError = string.Empty;
        private Func<Task>? postPasswordAction = null;

        protected override async Task OnInitializedAsync()
        {
            await LoadFilters();
        }

        protected override async Task OnParametersSetAsync()
        {
            invoiceType = InvoiceTypeString.ToLower() switch
            {
                "commercial" => 1,
                "bank" => 2,
                _ => 0
            };
            await LoadData();
        }

        public string GetPageTitle()
        {
            return invoiceType switch
            {
                1 => "Commercial Invoice List",
                2 => "Bank Invoice List",
                _ => "Custom Invoice List"
            };
        }

        private async Task LoadFilters()
        {
            companies = (await InvoiceListService.GetCompaniesAsync()).ToList();
            customers = (await InvoiceListService.GetCustomersAsync()).ToList();
            countries = (await InvoiceListService.GetCountriesAsync()).ToList();
        }

        private async Task LoadData()
        {
            isLoading = true;
            StateHasChanged();

            allItems = (await InvoiceListService.GetInvoicesAsync(
                invoiceType: invoiceType,
                companyId: selectedCompanyId,
                customerCode: selectedCustomer?.CustCode ?? "0",
                country: selectedCountry,
                dateRange: selectedDateRange,
                fromDate: fromDate,
                toDate: toDate,
                searchText: searchText
            )).ToList();

            isLoading = false;
            StateHasChanged();
        }

        private async Task<IEnumerable<CustomerLookupModel>> SearchCustomers(string searchText)
        {
            return await Task.FromResult(customers.Where(x => x.CustCode.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        private async Task OnCustomerChanged(CustomerLookupModel? customer)
        {
            selectedCustomer = customer;
            await LoadData();
        }

        private async Task OnCompanyChanged(ChangeEventArgs e)
        {
            if (long.TryParse(e.Value?.ToString(), out long compId))
            {
                selectedCompanyId = compId;
                await LoadData();
            }
        }

        private async Task OnCountryChanged(ChangeEventArgs e)
        {
            selectedCountry = e.Value?.ToString() ?? string.Empty;
            await LoadData();
        }

        private async Task OnDateRangeChanged()
        {
            if (selectedDateRange != "custom")
            {
                if (int.TryParse(selectedDateRange, out int days))
                {
                    fromDate = DateTime.Today.AddDays(-days);
                    toDate = DateTime.Today;
                }
                await LoadData();
            }
        }

        private async Task OpenRowContextMenu(MouseEventArgs e, InvoiceListModel item)
        {
            contextMenuRowItem = item;
            await BlazorContextMenuService.ShowMenu("rowContextMenu", (int)e.ClientX, (int)e.ClientY);
        }

        private async Task CheckPasswordThenExecute(string settingName, Func<Task> action)
        {
            passwordSettingName = settingName;
            enteredPassword = string.Empty;
            passwordError = string.Empty;
            postPasswordAction = action;
            showPasswordModal = true;
            await Task.CompletedTask;
        }

        private async Task VerifyPasswordAndExecute()
        {
            try
            {
                string correctPassword = await DbHelper.getDatabasePasswordAsync(passwordSettingName);
                if (enteredPassword == correctPassword)
                {
                    showPasswordModal = false;
                    if (postPasswordAction != null)
                    {
                        await postPasswordAction.Invoke();
                    }
                }
                else
                {
                    passwordError = "Incorrect password. Please try again.";
                }
            }
            catch (Exception ex)
            {
                passwordError = $"Error validating password: {ex.Message}";
            }
        }

        private async Task ShowOptionsMenu(MouseEventArgs e)
        {
            await BlazorContextMenuService.ShowMenu("optionsBarMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        private async Task PrintInvoiceListReport()
        {
            var f = fromDate ?? DateTime.Today.AddDays(-180);
            var t = toDate ?? DateTime.Today;
            string sel = $"{{VCustomInvoiceList.DT}} in Date({f.Year}, {f.Month}, {f.Day}) to Date({t.Year}, {t.Month}, {t.Day})";
            
            if (selectedCustomer != null && selectedCustomer.CustCode != "0")
            {
                sel += $" AND {{VCustomInvoiceList.CustCode}}='{selectedCustomer.CustCode}'";
            }
            if (!string.IsNullOrEmpty(selectedCountry))
            {
                sel += $" AND {{VCustomInvoiceList.Country}}='{selectedCountry}'";
            }

            string customersStr = $"{(selectedCustomer?.Name ?? "<All Customers>")} {(string.IsNullOrEmpty(selectedCountry) ? "" : selectedCountry)}";
            string dateRangeStr = $"{f:dd-MMM-yyyy} to {t:dd-MMM-yyyy}";

            var request = new ReportRequest
            {
                ReportName = "CustomInvoiceList.rpt",
                SelectionFormula = sel,
                FormulaValues = new Dictionary<string, object>
                {
                    { "Customers", $"'{customersStr}'" },
                    { "DateRange", $"'{dateRangeStr}'" },
                    { "bCustom", true }
                }
            };

            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task DeleteInvoice(ItemClickEventArgs e)
        {
            if (contextMenuRowItem == null) return;

            bool confirm = await JS.InvokeAsync<bool>("confirm", $"Are you sure you want to delete invoice {contextMenuRowItem.CustomInvoice}?");
            if (!confirm) return;

            _ = CheckPasswordThenExecute("DeleteExportEntries", async () =>
            {
                try
                {
                    await InvoiceListService.DeleteInvoiceAsync(contextMenuRowItem.CustomInvoice);
                    NotificationServiceManager.ShowSuccess("Success", "Invoice deleted successfully.");
                    await LoadData();
                }
                catch (Exception ex)
                {
                    NotificationServiceManager.ShowError("Delete Error", ex.Message);
                }
            });
        }

        private async Task PrintReport(string actionKey)
        {
            if (contextMenuRowItem == null) return;

            string invNo = contextMenuRowItem.CustomInvoice;
            string reportName = "";
            string selectionFormula = "";
            var formulaValues = new Dictionary<string, object>();
            bool isCustom = (invoiceType == 0);

            switch (actionKey)
            {
                case "rptInvoice":
                    if (invoiceType == 2)
                    {
                        reportName = "Bank_Invoice.rpt";
                    }
                    else if (invoiceType == 1)
                    {
                        reportName = "rptCommercialInvoice.rpt";
                    }
                    else
                    {
                        reportName = "rptCustomInvoice.rpt";
                    }
                    selectionFormula = $"{{VrptCustomInvoiceDetail.CustomInvoice}}='{invNo}'";
                    formulaValues["Discount"] = 0;
                    break;

                case "rptCustomInvoice": // Actual Invoice
                    reportName = "Print_Actual_Invoice.rpt";
                    selectionFormula = $"{{VrptCustomInvoiceDetail.CustomInvoice}}='{invNo}'";
                    formulaValues["Discount"] = 0;
                    break;

                case "rptCustomInvoice_USA":
                    reportName = isCustom ? "Invoice_USA.rpt" : "Commercial_Invoice_USA.rpt";
                    selectionFormula = $"{{VrptCustomInvoiceDetail.CustomInvoice}}='{invNo}'";
                    break;

                case "rptInvoiceasHandTools":
                    reportName = "Invoice_Handtools.rpt";
                    selectionFormula = $"{{VrptProformas.CustomInvoice}}='{invNo}'";
                    break;

                case "rptInvoiceWOCustomerAddress":
                    reportName = isCustom ? "rptSampleInvoiceWOAddress.rpt" : "Commercial_Invoice_WO_Address.rpt";
                    selectionFormula = $"{{VrptCustomInvoiceDetail.CustomInvoice}}='{invNo}'";
                    formulaValues["Discount"] = 0;
                    break;

                case "rptInvoiceDiscount":
                    reportName = isCustom ? "rptCustomInvoice.rpt" : "rptCommercialInvoice.rpt";
                    selectionFormula = $"{{VrptCustomInvoiceDetail.CustomInvoice}}='{invNo}'";
                    formulaValues["Discount"] = 0;
                    break;

                case "rptComInvoice_Pay":
                    reportName = "Payment_Invoice.rpt";
                    selectionFormula = $"{{VrptCustomInvoiceDetail.CustomInvoice}}='{invNo}'";
                    break;

                case "rptUSAPaymentInvoice":
                    reportName = "Commercial_USA_Payment_Invoice.rpt";
                    selectionFormula = $"{{VrptProformas.CustomInvoice}}='{invNo}'";
                    break;

                case "rptComInvoice_Docs":
                    reportName = "Print_Payment_Docs.rpt";
                    selectionFormula = $"{{VrptCustomInvoiceDetail.CustomInvoice}}='{invNo}'";
                    break;

                case "rptCustomPListDetailed": // Packing List
                    if (invoiceType == 2)
                    {
                        reportName = "Bank_Packing_List.rpt";
                    }
                    else if (invoiceType == 1)
                    {
                        reportName = "rptCommercialPList.rpt";
                    }
                    else
                    {
                        reportName = "rptCustomPList.rpt";
                    }
                    selectionFormula = $"{{VCustomPList.CustomInvoice}}='{invNo}'";
                    break;

                case "rptGSP":
                    reportName = "rpt_GSP_Invoice.rpt";
                    selectionFormula = $"{{VrptProformas.CustomInvoice}}='{invNo}'";
                    break;

                case "rptCustomPListCombined":
                    reportName = isCustom ? "rptCustomPListDetailed.rpt" : "Commercial_Packing_List_Detailed.rpt";
                    selectionFormula = $"{{VCustomPList.CustomInvoice}}='{invNo}'";
                    break;

                case "rptCustomPListCombinedCustom":
                    reportName = isCustom ? "rptCustomPListDetailedCustom.rpt" : "Commercial_Packing_List_Detailed_Custom.rpt";
                    selectionFormula = $"{{VCustomPList.CustomInvoice}}='{invNo}'";
                    break;

                case "rptBillOfExchange":
                    reportName = "rptBillOfExchange.rpt";
                    selectionFormula = $"{{VRptProformas.CustomInvoice}}='{invNo}'";
                    break;

                case "rptInvPListComparison":
                    reportName = "InvoicePackingListComparison.rpt";
                    selectionFormula = $"{{CustomInvoice.CustomInvoice}}='{invNo}'";
                    break;

                case "rptLabLetter":
                    reportName = "LaboratoryTestLetter.rpt";
                    selectionFormula = $"{{CustomInvoice.CustomInvoice}}='{invNo}'";
                    break;

                case "rptBankSlipLabTest":
                    reportName = "NationalBankSlip.rpt";
                    selectionFormula = $"{{CustomInvoice.CustomInvoice}}='{invNo}'";
                    break;

                case "rptUnderTaking":
                    reportName = "Undertaking.rpt";
                    selectionFormula = $"{{VCustomInvoice.CustomInvoice}}='{invNo}'";
                    break;

                case "rptLetterToBank":
                    reportName = "rptLetterToBank.rpt";
                    selectionFormula = $"{{VrptProformas.CustomInvoice}}='{invNo}'";
                    break;

                case "rptLetterToBankSimple":
                    reportName = "rptLetterToBankSimple.rpt";
                    selectionFormula = $"{{VrptProformas.CustomInvoice}}='{invNo}'";
                    break;

                case "rptBankSalesContract":
                    reportName = "Bank_Sales_Contract.rpt";
                    selectionFormula = $"{{VrptProformas.CustomInvoice}}='{invNo}'";
                    break;

                case "rptReasonForDelay":
                    reportName = "Bank_Reason_For_Delay.rpt";
                    selectionFormula = $"{{VrptProformas.CustomInvoice}}='{invNo}'";
                    break;

                case "rptReasonForDelaySBP":
                    reportName = "Bank_Reason_For_Delay_SBP.rpt";
                    selectionFormula = $"{{VrptProformas.CustomInvoice}}='{invNo}'";
                    break;

                case "rptLabReport":
                    reportName = "Bank_Lab_Report.rpt";
                    selectionFormula = $"{{VrptProformas.CustomInvoice}}='{invNo}'";
                    break;

                case "rptBankUnderTaking1":
                    reportName = "Bank_Undertaking1.rpt";
                    selectionFormula = $"{{VrptProformas.CustomInvoice}}='{invNo}'";
                    break;

                case "rptBankUnderTaking2":
                    reportName = "Bank_Undertaking2.rpt";
                    selectionFormula = $"{{VrptProformas.CustomInvoice}}='{invNo}'";
                    break;

                case "rptBankUnderTaking3":
                    reportName = "Bank_Undertaking3.rpt";
                    selectionFormula = $"{{VrptProformas.CustomInvoice}}='{invNo}'";
                    break;

                case "rptInvoiceExport":
                    reportName = "Commercial_Invoice_Export.rpt";
                    selectionFormula = $"{{VCustomInvoice.CustomInvoice}}='{invNo}'";
                    break;

                case "rptDeclarationOfConformity":
                    reportName = "Declartion_Of_Conformity.rpt";
                    selectionFormula = $"{{VCustomInvoice.CustomInvoice}}='{invNo}'";
                    break;

                case "rptCertificateOfCompliance":
                    reportName = "Certificate_Of_Compliance.rpt";
                    selectionFormula = $"{{VCustomInvoice.CustomInvoice}}='{invNo}'";
                    break;

                case "rptList":
                    await PrintInvoiceListReport();
                    return;

                default:
                    reportName = actionKey.EndsWith(".rpt", StringComparison.OrdinalIgnoreCase) ? actionKey : $"{actionKey}.rpt";
                    selectionFormula = $"{{VCustomInvoice.CustomInvoice}}='{invNo}'";
                    break;
            }

            var request = new ReportRequest
            {
                ReportName = reportName,
                SelectionFormula = selectionFormula,
                FormulaValues = formulaValues
            };

            await ReportNavigation.PrintReportAsync(request);
        }

        private void AddInvoice(ItemClickEventArgs e)
        {
            NavigationManager.NavigateTo("/export/custominvoices/new-custominvoice");
        }

        private void EditInvoice(ItemClickEventArgs e)
        {
            if (contextMenuRowItem == null) return;
            NavigationManager.NavigateTo($"/export/custominvoices/new-custominvoice/{contextMenuRowItem.CustomInvoice}");
        }

        private void UpdateRebateRates(ItemClickEventArgs e)
        {
            if (contextMenuRowItem == null) return;
            NotificationServiceManager.ShowWarning("Pending", "Update Rebate Rates functionality will be implemented soon.");
        }

        private void GenerateBatches(ItemClickEventArgs e)
        {
            if (contextMenuRowItem == null) return;
            NotificationServiceManager.ShowWarning("Pending", "Generate Batches functionality will be implemented soon.");
        }
    }
}
