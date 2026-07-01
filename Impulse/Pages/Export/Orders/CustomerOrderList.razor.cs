using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Constants;
using Impulse.Services;
using Impulse.Services.Export;

namespace Impulse.Pages.Export.Orders
{
    public partial class CustomerOrderList : ComponentBase
    {
        [Inject] private ICustomerOrderService CustomerOrderService { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private IDBHelperService DbHelper { get; set; } = null!;
        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; } = null!;

        [Parameter]
        [SupplyParameterFromQuery(Name = "viewType")]
        public int ViewType { get; set; } = 0; // 0 = Show All, 1 = Non Parts & Stock, 2 = Parts & Stock Only

        private bool isLoading = true;
        private string searchText = string.Empty;

        // Filters state
        private DateTime dtFrom = DateTime.Today.AddYears(-5);
        private DateTime dtTo = DateTime.Today;
        private bool filterByDeliveryDT = false;
        private string selectedOrderType = "<All Types>";
        private int selectedStatusFilter = 0; // default <Unshipped>

        // Lookups
        private List<CompanyLookupModel> companies = new();
        private List<CustomerLookupModel> customers = new();
        private List<string> countries = new();

        private CompanyLookupModel? selectedCompanyFilter = null;
        private CustomerLookupModel? selectedCustomerFilter = null;
        private string selectedCountryFilter = "<All Countries>";

        // List records
        private List<CustomerOrderListItemModel> allOrders = new();

        // Selected row state
        private CustomerOrderListItemModel? highlightedOrder = null;

        // Modal States
        private bool showPasswordModal = false;
        private string passwordSettingName = string.Empty;
        private string enteredPassword = string.Empty;
        private string passwordError = string.Empty;
        private Func<Task>? postPasswordAction = null;

        private bool showRemarksModal = false;
        private string remarksModalTitle = string.Empty;
        private string enteredRemarks = string.Empty;
        private Func<string, Task>? postRemarksAction = null;

        // Totals
        private decimal totalAmtPKR = 0;

        protected override async Task OnInitializedAsync()
        {
            await LoadLookups();
            await RefreshList();
        }

        protected override async Task OnParametersSetAsync()
        {
            await RefreshList();
        }

        private async Task LoadLookups()
        {
            try
            {
                var rawCompanies = await CustomerOrderService.GetCompaniesAsync();
                companies = rawCompanies.Select(c => new CompanyLookupModel 
                { 
                    EntryID = int.TryParse(c.DropDownValue_ID, out int val) ? val : 1, 
                    CompanyName = c.DropDownValue_Description 
                }).ToList();

                customers = await CustomerOrderService.GetCustomersAsync();
                await LoadCountries();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error Loading Lookups", ex.Message);
            }
        }

        private async Task LoadCountries()
        {
            try
            {
                if (selectedCustomerFilter != null && selectedCustomerFilter.CustCode != "0")
                {
                    countries = await CustomerOrderService.GetCountriesForCustomerAsync(selectedCustomerFilter.CustCode);
                }
                else
                {
                    countries = await CustomerOrderService.GetAllCountriesAsync();
                }
                selectedCountryFilter = "<All Countries>";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading countries: {ex.Message}");
            }
        }

        private async Task RefreshList()
        {
            isLoading = true;
            try
            {
                string custCode = selectedCustomerFilter?.CustCode ?? "0";
                int compRefId = selectedCompanyFilter?.EntryID ?? 0;

                allOrders = await CustomerOrderService.GetOrderListAsync(
                    dtFrom,
                    dtTo,
                    custCode,
                    selectedCountryFilter,
                    compRefId,
                    selectedOrderType,
                    selectedStatusFilter,
                    filterByDeliveryDT,
                    ViewType
                );

                // Recalculate totals
                totalAmtPKR = allOrders.Sum(o => o.OrderAmt * (o.ExchRate ?? 1));
                highlightedOrder = null;
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error Refreshing List", ex.Message);
            }
            finally
            {
                isLoading = false;
            }
        }

        private IEnumerable<CustomerOrderListItemModel> GetFilteredOrders()
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return allOrders;
            }
            return allOrders.Where(o => 
                o.OrderNo.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                o.InternalRefNo.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                o.CustCode.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                o.CompanyName.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                o.Country.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                o.OrderType.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                o.OrderStatus.Contains(searchText, StringComparison.OrdinalIgnoreCase)
            );
        }

        // ── Blazored Typeahead Filter Event Triggers ──

        public async Task<IEnumerable<CompanyLookupModel>> SearchCompanies(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText)) return companies;
            return await Task.FromResult(companies.Where(c => c.CompanyName.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        public async Task OnCompanyChanged(CompanyLookupModel? comp)
        {
            selectedCompanyFilter = comp;
            await RefreshList();
        }

        public async Task<IEnumerable<CustomerLookupModel>> SearchCustomers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText)) return customers;
            return await Task.FromResult(customers.Where(c => c.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        public async Task OnCustomerChanged(CustomerLookupModel? cust)
        {
            selectedCustomerFilter = cust;
            await LoadCountries();
            await RefreshList();
        }

        public async Task OnCountryChanged(ChangeEventArgs e)
        {
            selectedCountryFilter = e.Value?.ToString() ?? "<All Countries>";
            await RefreshList();
        }

        public async Task OnOrderTypeChanged(ChangeEventArgs e)
        {
            selectedOrderType = e.Value?.ToString() ?? "<All Types>";
            await RefreshList();
        }

        public async Task OnStatusFilterChanged(ChangeEventArgs e)
        {
            selectedStatusFilter = int.TryParse(e.Value?.ToString(), out int val) ? val : 0;
            await RefreshList();
        }

        public async Task OnFilterByDeliveryDTChanged(ChangeEventArgs e)
        {
            filterByDeliveryDT = (bool)(e.Value ?? false);
            await RefreshList();
        }

        private void HighlightOrder(CustomerOrderListItemModel order)
        {
            highlightedOrder = order;
        }

        private async Task OpenRowMenu(MouseEventArgs e, CustomerOrderListItemModel order)
        {
            highlightedOrder = order;
            await BlazorContextMenuService.ShowMenu("rowContextMenu", (int)e.ClientX, (int)e.ClientY + 15, order);
        }

        private async Task OpenOptionsBarMenu(MouseEventArgs e)
        {
            await BlazorContextMenuService.ShowMenu("optionsBarMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        // ── Password Protection Check Helper ──

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

        // ── Remarks Prompt Modal Helper ──

        private void PromptRemarks(string title, Func<string, Task> action)
        {
            remarksModalTitle = title;
            enteredRemarks = string.Empty;
            postRemarksAction = action;
            showRemarksModal = true;
        }

        private async Task SaveRemarks()
        {
            showRemarksModal = false;
            if (postRemarksAction != null)
            {
                await postRemarksAction.Invoke(enteredRemarks);
            }
        }

        // ── Row Context Menu Commands ──

        private void EditOrder(string orderNo)
        {
            if (ViewType == 0 && (selectedCompanyFilter?.CompanyName == "IAA" || selectedCompanyFilter?.CompanyName == "Kami"))
            {
                _ = CheckPasswordThenExecute("EditExportDocs", () =>
                {
                    NavigationManager.NavigateTo($"/export/new-customer-order/{orderNo}");
                    return Task.CompletedTask;
                });
            }
            else
            {
                NavigationManager.NavigateTo($"/export/new-customer-order/{orderNo}");
            }
        }

        private async Task CopyOrder(string orderNo)
        {
            try
            {
                var order = await CustomerOrderService.GetOrderAsync(orderNo);
                if (order != null)
                {
                    string newOrderNo = order.InternalRefNo + "-COPY";
                    order.OrderNo = newOrderNo;
                    order.InternalRefNo = await CustomerOrderService.GetNextInternalRefNoAsync(order.CompanyRefID);
                    order.DT = DateTime.Today;
                    foreach (var line in order.OrderItems)
                    {
                        line.ID = 0;
                    }
                    bool success = await CustomerOrderService.SaveOrderAsync(order, new List<int>());
                    if (success)
                    {
                        NotificationServiceManager.ShowSuccess("Success", $"Order duplicated as {newOrderNo}");
                        await RefreshList();
                    }
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Copy Failed", ex.Message);
            }
        }

        private void DeleteOrder(string orderNo)
        {
            _ = CheckPasswordThenExecute("Delete_Customer_Order", async () =>
            {
                try
                {
                    bool success = await CustomerOrderService.DeleteCustomerOrderAsync(orderNo);
                    if (success)
                    {
                        NotificationServiceManager.ShowSuccess("Success", $"Order {orderNo} deleted.");
                        await RefreshList();
                    }
                    else
                    {
                        NotificationServiceManager.ShowError("Failed", "Cannot delete order: Proforma invoice exists.");
                    }
                }
                catch (Exception ex)
                {
                    NotificationServiceManager.ShowError("Delete Failed", ex.Message);
                }
            });
        }

        private void MarkAsShipped(CustomerOrderListItemModel order)
        {
            _ = UpdateFinalOrderStatus(order.OrderNo, order.CustCode, order.Country, 0, "Marked as Shipped");
        }

        private void CancelOrder(CustomerOrderListItemModel order)
        {
            _ = CheckPasswordThenExecute("DeleteExportEntries", async () =>
            {
                await UpdateFinalOrderStatus(order.OrderNo, order.CustCode, order.Country, 1, "Cancelled");
            });
        }

        private void HoldOrder(CustomerOrderListItemModel order)
        {
            PromptRemarks("Enter Hold Remarks", async (remarks) =>
            {
                await UpdateFinalOrderStatus(order.OrderNo, order.CustCode, order.Country, 2, remarks);
            });
        }

        private async Task UnCancelUnHoldOrder(string orderNo)
        {
            try
            {
                bool success = await CustomerOrderService.DeleteOrderFinalStatusAsync(orderNo);
                if (success)
                {
                    NotificationServiceManager.ShowSuccess("Success", "Order status restored.");
                    await RefreshList();
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Update Failed", ex.Message);
            }
        }

        private async Task UpdateFinalOrderStatus(string orderNo, string custCode, string country, int cancelledCode, string remarks)
        {
            try
            {
                string user = "BlazorUser";
                string machine = "BlazorWeb";

                bool success = await CustomerOrderService.UpdateOrderFinalStatusAsync(orderNo, custCode, country, cancelledCode, remarks, user, machine);
                if (success)
                {
                    NotificationServiceManager.ShowSuccess("Success", $"Order status updated.");
                    await RefreshList();
                }
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Update Failed", ex.Message);
            }
        }

        // ── Crystal Reports Triggering ──

        private async Task PrintOrderWrangler(string orderNo, bool printBalance = false, bool makerWise = false, bool stockSummary = false, bool groupwise = false, bool markingPlan = false, bool withValue = false, bool productionOrder = false, bool orderStatus = false, bool orderDetails = false, int reportType = 0)
        {
            try
            {
                string rptName = ReportNames.Export.ProdOrderWrangler;
                string selectionFormula = $"{{CustOrders.OrderNo}}='{orderNo}'";

                if (groupwise)
                {
                    rptName = ReportNames.Export.ProdOrderGroupWise;
                }
                else if (markingPlan)
                {
                    rptName = ReportNames.Export.ProdOrderMarkingPlan;
                }
                else if (withValue)
                {
                    rptName = ReportNames.Export.ProdOrderWValue;
                }
                else if (productionOrder)
                {
                    rptName = ReportNames.Export.ProdOrderWrangler1;
                }
                else if (orderStatus)
                {
                    if (reportType == 1) rptName = ReportNames.Export.ProdOrderRM;
                    else if (reportType == 2) rptName = ReportNames.Export.ProdOrderPacking;
                    else rptName = ReportNames.Export.ProdOrderStatus;
                }
                else if (orderDetails)
                {
                    rptName = ReportNames.Export.ProdOrderDetails;
                }
                else if (printBalance)
                {
                    rptName = ReportNames.Export.ProdOrderBalance;
                }
                else
                {
                    if (reportType == 99) rptName = ReportNames.Export.PrintOrderSlip;
                    else if (reportType == 98) rptName = ReportNames.Export.PrintOrderPictures;
                }

                var parameters = new Dictionary<string, object>
                {
                    { "@OrderNo", orderNo }
                };

                var formulaValues = new Dictionary<string, object>
                {
                    { "UserName", "BlazorUser" }
                };

                var request = new ReportRequest
                {
                    ReportName = rptName,
                    SelectionFormula = selectionFormula,
                    Parameters = parameters,
                    FormulaValues = formulaValues
                };

                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Report Failed", ex.Message);
            }
        }

        private async Task PrintMaterialConsumptionSheet(string orderNo, bool articleWise = false)
        {
            string reportName = articleWise ? "MaterialConsumptionSheet_Articlewise.rpt" : "MaterialConsumptionSheet.rpt";
            var request = new ReportRequest
            {
                ReportName = reportName,
                SelectionFormula = $"{{VrptOrders.OrderNo}}='{orderNo}'"
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintCustomerPO(string orderNo)
        {
            var request = new ReportRequest
            {
                ReportName = ReportNames.Export.CustomerPO,
                SelectionFormula = $"{{FCustomerOrders.OrderNo}}='{orderNo}'"
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintDeliverySchedule(string orderNo)
        {
            var request = new ReportRequest
            {
                ReportName = ReportNames.Export.OrderItemList,
                FormulaValues = new Dictionary<string, object>
                {
                    { "Filters", $"Order No.:{orderNo}" }
                }
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintScoreBoard(string orderNo)
        {
            var request = new ReportRequest
            {
                ReportName = ReportNames.Export.ScoreBoard,
                Parameters = new Dictionary<string, object>
                {
                    { "@OrderNo", orderNo }
                }
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintOrderBalanceReportCT(string orderNo)
        {
            var request = new ReportRequest
            {
                ReportName = ReportNames.Export.BalanceOrdersCT,
                SelectionFormula = $"{{VFOrderItemswithShippedQty.DT}} >= #{dtFrom:yyyy-MM-dd}# AND {{VFOrderItemswithShippedQty.DT}} <= #{dtTo:yyyy-MM-dd}# AND {{VFOrderItemswithShippedQty.OrderNo}}='{orderNo}'"
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintOrderList(bool balanceListOnly = false)
        {
            var formulaValues = new Dictionary<string, object>
            {
                { "ForCustomer", $"{selectedCustomerFilter?.Name ?? "<All Customers>"} {selectedCountryFilter}" },
                { "ForFromTo", $"{dtFrom:dd-MMM-yyyy} to {dtTo:dd-MMM-yyyy}" },
                { "OrderType", selectedOrderType }
            };

            var request = new ReportRequest
            {
                ReportName = ReportNames.Export.FOrderList,
                SelectionFormula = string.Empty,
                FormulaValues = formulaValues
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintCombinedOrderReport(int mode)
        {
            string rptName = ReportNames.Export.PartShippedItemwise;
            if (mode == 1) rptName = ReportNames.Export.CombinedOrderItemwise;
            else if (mode == 2) rptName = ReportNames.Export.PartShippedOrders;

            string cond = $"{{FCustomerOrders.DT}} >= #{dtFrom:yyyy-MM-dd}# AND {{FCustomerOrders.DT}} <= #{dtTo:yyyy-MM-dd}#";
            if (selectedCustomerFilter != null && selectedCustomerFilter.CustCode != "0")
            {
                cond += $" AND {{FCustomerOrders.CustCode}}='{selectedCustomerFilter.CustCode}'";
            }
            if (!string.IsNullOrEmpty(selectedCountryFilter) && selectedCountryFilter != "<All Countries>")
            {
                cond += $" AND {{FCustomerOrders.Country}}='{selectedCountryFilter}'";
            }

            var request = new ReportRequest
            {
                ReportName = rptName,
                SelectionFormula = cond
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintProductionExportCost(string orderNo)
        {
            var request = new ReportRequest
            {
                ReportName = ReportNames.Export.ProductionExportCost,
                SelectionFormula = $"{{VProductionExportCost.OrderNo}}='{orderNo}'"
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintOrderStatusLogReport(bool extendedLog)
        {
            var request = new ReportRequest
            {
                ReportName = extendedLog ? ReportNames.Export.OrderStatusLogEx : ReportNames.Export.OrderStatusLog,
                Parameters = new Dictionary<string, object>
                {
                    { "@CustCode", selectedCustomerFilter?.CustCode == "0" ? "" : (selectedCustomerFilter?.CustCode ?? "") },
                    { "@Country", selectedCountryFilter == "<All Countries>" ? "" : selectedCountryFilter },
                    { "@ShipStatus", selectedStatusFilter },
                    { "@DTFrom", dtFrom },
                    { "@DTTo", dtTo }
                }
            };
            if (extendedLog)
            {
                request.Parameters.Add("@DeliveryDTFilter", filterByDeliveryDT);
            }
            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintPlanningPurchaseLogReport()
        {
            var request = new ReportRequest
            {
                ReportName = ReportNames.Export.OrderPurchasePlanningLog,
                Parameters = new Dictionary<string, object>
                {
                    { "@CustCode", selectedCustomerFilter?.CustCode == "0" ? "" : (selectedCustomerFilter?.CustCode ?? "") },
                    { "@Country", selectedCountryFilter == "<All Countries>" ? "" : selectedCountryFilter },
                    { "@ShipStatus", selectedStatusFilter },
                    { "@DTFrom", dtFrom },
                    { "@DTTo", dtTo }
                }
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintProductionTimeline()
        {
            var request = new ReportRequest
            {
                ReportName = ReportNames.Export.ProductionTimeLine,
                Parameters = new Dictionary<string, object>
                {
                    { "@CustCode", selectedCustomerFilter?.CustCode == "0" ? "" : (selectedCustomerFilter?.CustCode ?? "") },
                    { "@Country", selectedCountryFilter == "<All Countries>" ? "" : selectedCountryFilter },
                    { "@ShipStatus", selectedStatusFilter },
                    { "@DTFrom", dtFrom },
                    { "@DTTo", dtTo }
                }
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintUrgentRequestReport()
        {
            string cond = $"{{VrptOrders.DT}} >= #{dtFrom:yyyy-MM-dd}# AND {{VrptOrders.DT}} <= #{dtTo:yyyy-MM-dd}# AND {{VrptOrders.DeliveryStatus}}=2";
            if (selectedCustomerFilter != null && selectedCustomerFilter.CustCode != "0")
            {
                cond += $" AND {{VrptOrders.CustCode}}='{selectedCustomerFilter.CustCode}'";
            }
            if (selectedOrderType != "<All Types>")
            {
                cond += $" AND {{VrptOrders.OrderType}}='{selectedOrderType}'";
            }

            var request = new ReportRequest
            {
                ReportName = ReportNames.Export.UrgentRequest,
                SelectionFormula = cond
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintOrdersLogSheetReport()
        {
            var request = new ReportRequest
            {
                ReportName = ReportNames.Export.OrdersLogSheet,
                Parameters = new Dictionary<string, object>
                {
                    { "@CustCode", selectedCustomerFilter?.CustCode == "0" ? "" : (selectedCustomerFilter?.CustCode ?? "") },
                    { "@Country", selectedCountryFilter == "<All Countries>" ? "" : selectedCountryFilter },
                    { "@ShipStatus", selectedStatusFilter },
                    { "@DTFrom", dtFrom },
                    { "@DTTo", dtTo },
                    { "@DeliveryDTFilter", filterByDeliveryDT }
                }
            };
            await ReportNavigation.PrintReportAsync(request);
        }
    }
}
