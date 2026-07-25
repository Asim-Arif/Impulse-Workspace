using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Stock.RMIssuance
{
    public partial class RMIssuance : ComponentBase
    {
        [Inject] private Services.Stock.IRMIssuanceService RMIssuanceService { get; set; }
        [Inject] private Services.IAuditService AuditService { get; set; }
        [Inject] private Services.INotificationService NotificationService { get; set; }
        [Inject] private Services.IReportNavigationService ReportNavigation { get; set; }
        [Inject] private NavigationManager NavigationManager { get; set; }

        // State
        private bool ShowPicture { get; set; }
        private bool IsEmployee { get; set; } = true; // Toggle state
        private bool PrintWhenSaved { get; set; } = false;
        private bool IsSaving { get; set; } = false;

        private RMIssuanceSaveRequest Request { get; set; } = new() { DT = DateTime.Today };

        // Lookups
        private List<EmployeeLookupModel> AllEmployees = new();
        private List<MakerLookupModel> AllMakers = new();
        private List<DataAccessLibrary.Models.ViewModels.Stock.RMLookupModel> AllMaterials = new();
        private List<RMLocationModel> AvailableLocations = new();
        private List<RMHistoryRowModel> IssuanceHistory = new();

        // Selections
        private EmployeeLookupModel? SelectedEmployee;
        private MakerLookupModel? _selectedMaker;
        private MakerLookupModel? SelectedMaker
        {
            get => _selectedMaker;
            set
            {
                _selectedMaker = value;
                SelectedProductionIssuance = null;
                _ = LoadHistoryAsync();
            }
        }
        private EmployeeLookupModel? AuthorizedBy;
        
        private IssuanceReceiptLookupModel? _selectedProductionIssuance;
        private IssuanceReceiptLookupModel? SelectedProductionIssuance
        {
            get => _selectedProductionIssuance;
            set
            {
                _selectedProductionIssuance = value;
                Request.VI_RefID = value?.EntryID ?? 0;
                SelectedMaterial = null; // Clear material as the allowed list has changed
            }
        }

        private DataAccessLibrary.Models.ViewModels.Stock.RMLookupModel? _selectedMaterial;
        private DataAccessLibrary.Models.ViewModels.Stock.RMLookupModel? SelectedMaterial 
        { 
            get => _selectedMaterial; 
            set 
            {
                _selectedMaterial = value;
                _ = OnMaterialSelected();
            }
        }
        
        private RMLocationModel? SelectedLocation;
        private double QtyToIssue;
        private string CurrentUnit = "Qty";
        private double CurrentPrice = 0;
        private double TotalQtyInStock = 0;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                var empTask = RMIssuanceService.GetEmployeesAsync();
                var makerTask = RMIssuanceService.GetMakersAsync();
                var matTask = RMIssuanceService.GetMaterialsAsync();

                await Task.WhenAll(empTask, makerTask, matTask);

                AllEmployees = empTask.Result.ToList();
                AllMakers = makerTask.Result.ToList();
                AllMaterials = matTask.Result.ToList();

                Request.UserName = AuditService.GetCurrentUserName();
                Request.MachineName = AuditService.GetClientIpAddress();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Load Error", ex.Message);
            }
        }

        // -------------------------------------------------------------------------------------
        // Typeahead Searches
        // -------------------------------------------------------------------------------------
        private Task<IEnumerable<EmployeeLookupModel>> SearchEmployees(string searchText)
        {
            var results = string.IsNullOrWhiteSpace(searchText) 
                ? AllEmployees 
                : AllEmployees.Where(x => x.Display.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList();
            return Task.FromResult<IEnumerable<EmployeeLookupModel>>(results);
        }

        private Task<IEnumerable<MakerLookupModel>> SearchMakers(string searchText)
        {
            var results = string.IsNullOrWhiteSpace(searchText) 
                ? AllMakers 
                : AllMakers.Where(x => x.VenderName.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList();
            return Task.FromResult<IEnumerable<MakerLookupModel>>(results);
        }

        private async Task<IEnumerable<DataAccessLibrary.Models.ViewModels.Stock.RMLookupModel>> SearchMaterials(string searchText)
        {
            List<DataAccessLibrary.Models.ViewModels.Stock.RMLookupModel> source = AllMaterials;
            if (SelectedProductionIssuance != null)
            {
                source = (await RMIssuanceService.GetMaterialsByIssuanceAsync(SelectedProductionIssuance.EntryID)).ToList();
            }

            if (string.IsNullOrWhiteSpace(searchText))
                return source.Take(50);

            return source.Where(m => 
                (m.RMName != null && m.RMName.Contains(searchText, StringComparison.OrdinalIgnoreCase)) || 
                (m.RMID1 != null && m.RMID1.Contains(searchText, StringComparison.OrdinalIgnoreCase)))
                .Take(50);
        }

        private async Task<IEnumerable<IssuanceReceiptLookupModel>> SearchProductionIssuances(string searchText)
        {
            if (SelectedMaker == null) return Array.Empty<IssuanceReceiptLookupModel>();
            var issuances = await RMIssuanceService.GetProductionIssuancesAsync(SelectedMaker.VendID);
            
            if (string.IsNullOrWhiteSpace(searchText)) return issuances.Take(50);
            
            return issuances.Where(i => i.Display.Contains(searchText, StringComparison.OrdinalIgnoreCase)).Take(50);
        }

        private Task<IEnumerable<RMLocationModel>> SearchLocations(string searchText)
        {
            var results = string.IsNullOrWhiteSpace(searchText) 
                ? AvailableLocations 
                : AvailableLocations.Where(x => x.Display.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList();
            return Task.FromResult<IEnumerable<RMLocationModel>>(results);
        }

        // -------------------------------------------------------------------------------------
        // Actions
        // -------------------------------------------------------------------------------------
        private async Task OnMaterialSelected()
        {
            SelectedLocation = null;
            AvailableLocations.Clear();
            QtyToIssue = 0;
            CurrentUnit = "Qty";
            CurrentPrice = 0;
            TotalQtyInStock = 0;

            if (SelectedMaterial != null)
            {
                try
                {
                    var details = await RMIssuanceService.GetRMDetailsAsync(SelectedMaterial.RMID);
                    if (details != null)
                    {
                        CurrentUnit = details.Unit;
                        CurrentPrice = details.Rate ?? 0;
                        TotalQtyInStock = details.QtyInStock ?? 0;
                    }
                    var locs = await RMIssuanceService.GetMaterialLocationsAsync(SelectedMaterial.RMID);
                    AvailableLocations = locs.ToList();
                }
                catch (Exception ex)
                {
                    NotificationService.ShowError("Error", $"Failed to load material details: {ex.Message}");
                }
            }
            StateHasChanged();
        }

        private void HandleQtyKeyUp(KeyboardEventArgs e)
        {
            if (e.Key == "Enter" && SelectedLocation != null && QtyToIssue > 0)
            {
                AddItem();
            }
        }

        private void AddItem()
        {
            if (SelectedMaterial == null || SelectedLocation == null || QtyToIssue <= 0) return;

            if (QtyToIssue > SelectedLocation.ShelfQty)
            {
                NotificationService.ShowWarning("Insufficient Stock", $"Cannot issue {QtyToIssue}. Only {SelectedLocation.ShelfQty} available in this location.");
                return;
            }
            
            if (QtyToIssue > TotalQtyInStock)
            {
                NotificationService.ShowWarning("Insufficient Stock", $"Cannot issue {QtyToIssue}. Only {TotalQtyInStock} available across all stores.");
                return;
            }

            var newRow = new RMIssuanceDetailRow
            {
                RMID = SelectedMaterial.RMID,
                RMID1 = SelectedMaterial.RMID1,
                RMName = SelectedMaterial.RMName,
                RMSize = SelectedMaterial.RMSize,
                SizeUnit = SelectedMaterial.SizeUnit,
                Unit = CurrentUnit,
                Store = SelectedLocation.Store,
                Location = SelectedLocation.Display,
                Shelf_RefID = SelectedLocation.Shelf_RefID,
                Qty = QtyToIssue,
                Rate = CurrentPrice,
                BatchNo = SelectedLocation.BatchNo,
                LotNo = SelectedLocation.LotNo,
                TotalQtyInStock = TotalQtyInStock
            };

            Request.Rows.Add(newRow);

            // Deduct locally from total and location so user can't double-issue beyond capacity
            SelectedLocation.ShelfQty -= QtyToIssue;
            TotalQtyInStock -= QtyToIssue;

            // Reset selection for next item
            SelectedMaterial = null; // triggers OnMaterialSelected reset
            
            _ = LoadHistoryAsync(); // Fire-and-forget history load
        }

        private void RemoveItem(RMIssuanceDetailRow row)
        {
            Request.Rows.Remove(row);
        }

        private async Task LoadHistoryAsync()
        {
            if (IsEmployee && SelectedEmployee != null)
            {
                IssuanceHistory = (await RMIssuanceService.GetIssuanceHistoryAsync(false, SelectedEmployee.EmpID, DateTime.Today.AddMonths(-2), DateTime.Today)).ToList();
            }
            else if (!IsEmployee && SelectedMaker != null)
            {
                IssuanceHistory = (await RMIssuanceService.GetIssuanceHistoryAsync(true, SelectedMaker.VendID.ToString(), DateTime.Today.AddMonths(-2), DateTime.Today)).ToList();
            }
            else
            {
                IssuanceHistory.Clear();
            }
            StateHasChanged();
        }

        private void Close()
        {
            NavigationManager.NavigateTo("/stock");
        }

        private async Task SaveAsync(bool saveAndNew)
        {
            if (!Request.Rows.Any())
            {
                NotificationService.ShowWarning("Validation", "No items added to issue.");
                return;
            }

            if (IsEmployee)
            {
                if (SelectedEmployee == null)
                {
                    NotificationService.ShowWarning("Validation", "Please select an Employee.");
                    return;
                }
                Request.IssdToEmpID = SelectedEmployee.Display;
                Request.MakerVendID = 0;
            }
            else
            {
                if (SelectedMaker == null)
                {
                    NotificationService.ShowWarning("Validation", "Please select a Vendor.");
                    return;
                }
                Request.MakerVendID = SelectedMaker.VendID;
                Request.IssdToEmpID = ""; 
            }

            if (AuthorizedBy == null)
            {
                NotificationService.ShowWarning("Validation", "Please select Authorized By.");
                return;
            }
            Request.AuthorizedByEmpID = AuthorizedBy.Display;

            IsSaving = true;
            StateHasChanged();

            try
            {
                string newIssNo = await RMIssuanceService.SaveIssuanceTransactionAsync(Request);

                NotificationService.ShowSuccess("Saved", $"Issuance Note #{newIssNo} saved successfully.");

                if (PrintWhenSaved)
                {
                    await ReportNavigation.PrintReportAsync(new Services.ReportRequest
                    {
                        ReportName = "IssuanceNote.rpt",
                        SelectionFormula = $"{{IssItemsSimple.IssNo}}='{newIssNo}'"
                    });
                }

                if (saveAndNew)
                {
                    // Reset
                    Request = new RMIssuanceSaveRequest { DT = DateTime.Today, UserName = Request.UserName, MachineName = Request.MachineName };
                    SelectedEmployee = null;
                    SelectedMaker = null;
                    SelectedProductionIssuance = null;
                    AuthorizedBy = null;
                    IssuanceHistory.Clear();
                }
                else
                {
                    Close();
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Save Failed", ex.Message);
            }
            finally
            {
                IsSaving = false;
                StateHasChanged();
            }
        }
    }
}
