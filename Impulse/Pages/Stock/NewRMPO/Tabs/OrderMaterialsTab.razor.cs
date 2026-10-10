using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
using Impulse.Services;
using Impulse.Services.Stock;
using Microsoft.AspNetCore.Components;
using Radzen;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Stock.NewRMPO.Tabs
{
    public partial class OrderMaterialsTab : ComponentBase
    {
        [CascadingParameter]
        public VendOrderViewModel Order { get; set; }

        [Inject]
        private INewVendorService VendorService { get; set; }

        [Inject]
        private INewRMService RMService { get; set; }

        [Inject]
        private IDBHelperService DbHelper { get; set; }

        [Inject]
        private NotificationService NotificationService { get; set; }

        private List<NewVendorViewModel> Vendors = new();
        private List<RawMaterialLookupModel> AllMaterials = new();
        private List<EmployeeLookupDto> Employees = new();
        private List<PaymentTermLookupDto> PaymentTermsList = new();
        private List<DepartmentLookupDto> DepartmentsList = new();
        private List<CustomerOrderLookupDto> CustomerOrdersList = new();

        private NewVendorViewModel _selectedVendor;
        public NewVendorViewModel SelectedVendor
        {
            get => _selectedVendor;
            set
            {
                _selectedVendor = value;
                if (Order != null)
                {
                    Order.VendID = value?.AccNo;
                }
                _ = LoadPreviousRateAsync();
            }
        }

        private EmployeeLookupDto _selectedAttnPerson;
        public EmployeeLookupDto SelectedAttnPerson
        {
            get => _selectedAttnPerson;
            set
            {
                _selectedAttnPerson = value;
                if (Order != null)
                {
                    Order.AttnPerson = value?.Name ?? string.Empty;
                }
            }
        }

        private PaymentTermLookupDto _selectedPaymentTerm;
        public PaymentTermLookupDto SelectedPaymentTerm
        {
            get => _selectedPaymentTerm;
            set
            {
                _selectedPaymentTerm = value;
                if (Order != null)
                {
                    Order.PaymentTerms = value?.PaymentTerms ?? string.Empty;
                }
            }
        }

        private DepartmentLookupDto _selectedDepartment;
        public DepartmentLookupDto SelectedDepartment
        {
            get => _selectedDepartment;
            set
            {
                _selectedDepartment = value;
                if (Order != null)
                {
                    Order.Department = value?.Name ?? string.Empty;
                }
            }
        }

        private CustomerOrderLookupDto _selectedCustomerOrder;
        public CustomerOrderLookupDto SelectedCustomerOrder
        {
            get => _selectedCustomerOrder;
            set
            {
                _selectedCustomerOrder = value;
                if (Order != null)
                {
                    Order.CustomerPONo = value?.OrderNo ?? string.Empty;
                }
            }
        }

        private RawMaterialLookupModel _selectedMaterial;
        public RawMaterialLookupModel SelectedMaterial
        {
            get => _selectedMaterial;
            set
            {
                _selectedMaterial = value;
                if (value != null) SelectedMaterialId = value.RMID1;
                _ = LoadPreviousRateAsync();
            }
        }
        
        public string SelectedMaterialId { get; set; }
        public bool ShowPicture { get; set; }

        public RawMaterialLookupModel SelectedForMaterial { get; set; }
        public float? PreviousRate { get; set; }
        public float NewRate { get; set; }
        public float NewQty { get; set; }
        public string NewItemRemarks { get; set; }

        private int editingRowIndex = -1;

        private async Task LoadPreviousRateAsync()
        {
            if (SelectedMaterial == null || string.IsNullOrWhiteSpace(SelectedMaterial.RMID1) ||
                SelectedVendor == null || string.IsNullOrWhiteSpace(SelectedVendor.AccNo))
            {
                PreviousRate = null;
                return;
            }

            try
            {
                var vendId = SelectedVendor.AccNo.Replace("'", "''");
                var matId = SelectedMaterial.RMID1.Replace("'", "''");

                var rateResult = await DbHelper.getSingleDoubleValueasync(
                    "TOP 1 d.Rate",
                    "VendOrderDetail d INNER JOIN VendOrders o ON d.OrderNo = o.OrderNo",
                    $"WHERE o.VendID = '{vendId}' AND d.RMID = '{matId}' ORDER BY o.DT DESC, d.EntryID DESC");

                if (rateResult > 0)
                {
                    PreviousRate = (float)rateResult;
                    NewRate = (float)rateResult;
                }
                else
                {
                    PreviousRate = null;
                }
            }
            catch
            {
                PreviousRate = null;
            }
            await InvokeAsync(StateHasChanged);
        }

        protected override async Task OnInitializedAsync()
        {
            var vendorsResult = await VendorService.GetVendorsAsync();
            Vendors = vendorsResult != null ? vendorsResult : new List<NewVendorViewModel>();
            
            var materialsResult = await RMService.GetAvailableRawMaterialsAsync();
            AllMaterials = materialsResult != null ? materialsResult : new List<RawMaterialLookupModel>();

            try
            {
                Employees = await DbHelper.GetListAsync<EmployeeLookupDto>("empid AS EmpID, name AS Name, deptid AS DeptID", "Employees", "WHERE ISNULL(Active, 1) = 1 ORDER BY name");
            }
            catch { Employees = new(); }

            try
            {
                PaymentTermsList = await DbHelper.GetListAsync<PaymentTermLookupDto>("PaymentTerms", "RM_PO_Payment_Terms", "ORDER BY PaymentTerms");
            }
            catch { PaymentTermsList = new(); }

            try
            {
                DepartmentsList = await DbHelper.GetListAsync<DepartmentLookupDto>("name AS Name, Description", "Departments", "WHERE Active = 1 ORDER BY name");
            }
            catch { DepartmentsList = new(); }

            try
            {
                CustomerOrdersList = await DbHelper.GetListAsync<CustomerOrderLookupDto>("DISTINCT OrderNo", "FCustomerOrders", "WHERE OrderNo IS NOT NULL AND LTRIM(RTRIM(OrderNo)) <> '' ORDER BY OrderNo DESC");
            }
            catch { CustomerOrdersList = new(); }

            BindInitialSelections();
        }

        protected override void OnParametersSet()
        {
            BindInitialSelections();
        }

        private void BindInitialSelections()
        {
            if (Order == null) return;

            if (!string.IsNullOrEmpty(Order.VendID) && (_selectedVendor == null || _selectedVendor.AccNo != Order.VendID))
            {
                _selectedVendor = Vendors.FirstOrDefault(v => v.AccNo == Order.VendID);
            }

            if (!string.IsNullOrEmpty(Order.AttnPerson) && (_selectedAttnPerson == null || _selectedAttnPerson.Name != Order.AttnPerson))
            {
                _selectedAttnPerson = Employees.FirstOrDefault(e => e.Name == Order.AttnPerson || e.EmpID == Order.AttnPerson);
            }

            if (!string.IsNullOrEmpty(Order.PaymentTerms) && (_selectedPaymentTerm == null || _selectedPaymentTerm.PaymentTerms != Order.PaymentTerms))
            {
                _selectedPaymentTerm = PaymentTermsList.FirstOrDefault(p => p.PaymentTerms == Order.PaymentTerms);
            }

            if (!string.IsNullOrEmpty(Order.Department) && (_selectedDepartment == null || _selectedDepartment.Name != Order.Department))
            {
                _selectedDepartment = DepartmentsList.FirstOrDefault(d => d.Name == Order.Department);
            }

            if (!string.IsNullOrEmpty(Order.CustomerPONo) && (_selectedCustomerOrder == null || _selectedCustomerOrder.OrderNo != Order.CustomerPONo))
            {
                _selectedCustomerOrder = CustomerOrdersList.FirstOrDefault(c => c.OrderNo == Order.CustomerPONo);
            }
        }

        private async Task<IEnumerable<NewVendorViewModel>> SearchVendors(string searchText)
        {
            await Task.Delay(5);
            if (string.IsNullOrEmpty(searchText))
                return Vendors;

            return Vendors.Where(v => v.VendorName.Contains(searchText, System.StringComparison.OrdinalIgnoreCase)).ToList();
        }

        private async Task<IEnumerable<EmployeeLookupDto>> SearchEmployees(string searchText)
        {
            await Task.Delay(5);
            if (string.IsNullOrWhiteSpace(searchText))
                return Employees;

            return Employees.Where(e =>
                (e.Name != null && e.Name.Contains(searchText, System.StringComparison.OrdinalIgnoreCase)) ||
                (e.EmpID != null && e.EmpID.Contains(searchText, System.StringComparison.OrdinalIgnoreCase))
            ).ToList();
        }

        private async Task<IEnumerable<PaymentTermLookupDto>> SearchPaymentTerms(string searchText)
        {
            await Task.Delay(5);
            if (string.IsNullOrWhiteSpace(searchText))
                return PaymentTermsList;

            return PaymentTermsList.Where(p => p.PaymentTerms != null && p.PaymentTerms.Contains(searchText, System.StringComparison.OrdinalIgnoreCase)).ToList();
        }

        private async Task<IEnumerable<DepartmentLookupDto>> SearchDepartments(string searchText)
        {
            await Task.Delay(5);
            if (string.IsNullOrWhiteSpace(searchText))
                return DepartmentsList;

            return DepartmentsList.Where(d => d.Name != null && d.Name.Contains(searchText, System.StringComparison.OrdinalIgnoreCase)).ToList();
        }

        private async Task<IEnumerable<CustomerOrderLookupDto>> SearchCustomerOrders(string searchText)
        {
            await Task.Delay(5);
            if (string.IsNullOrWhiteSpace(searchText))
                return CustomerOrdersList;

            return CustomerOrdersList.Where(c => c.OrderNo != null && c.OrderNo.Contains(searchText, System.StringComparison.OrdinalIgnoreCase)).ToList();
        }

        private async Task<IEnumerable<RawMaterialLookupModel>> SearchMaterials(string searchText)
        {
            await Task.Delay(50);
            var query = AllMaterials.AsEnumerable();

            if (Order.POType == 3) // Male Grinding
            {
                // Filter specifically if needed
            }

            if (string.IsNullOrWhiteSpace(searchText))
                return query.Take(100).ToList();

            return query.Where(m => 
                (m.RMName != null && m.RMName.Contains(searchText, System.StringComparison.OrdinalIgnoreCase)) ||
                (m.RMID1 != null && m.RMID1.Contains(searchText, System.StringComparison.OrdinalIgnoreCase))
            ).Take(100).ToList();
        }

        private void OnPOTypeChanged()
        {
            // Reset selected items if needed
            SelectedMaterial = null;
            SelectedForMaterial = null;
        }

        private void AddMaterial()
        {
            if (SelectedMaterial == null)
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation Error", "Please Select The Material Which You Want To Order!");
                return;
            }

            if (NewQty <= 0)
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation Error", "Quantity Could Not Be Zero!");
                return;
            }

            if (Order.OrderDetails.Any(d => d.RMID == SelectedMaterial.RMID1))
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation Error", "This Material is Already Added to List!");
                return;
            }

            if (Order.POType > 0 && SelectedForMaterial == null)
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Validation Error", "Please Select For Material.");
                return;
            }

            var detail = new VendOrderDetailViewModel
            {
                RMID = SelectedMaterial.RMID1,
                RMName = SelectedMaterial.RMName,
                Unit = SelectedMaterial.Unit,
                Rate = NewRate,
                Qty = NewQty,
                ItemRemarks = NewItemRemarks,
                ForRMID = SelectedForMaterial != null && int.TryParse(SelectedForMaterial.RMID1, out int forRmId) ? forRmId : (int?)null,
                CCItem = false
            };

            Order.OrderDetails.Add(detail);

            // Set the vendor ID on master if selected
            if (SelectedVendor != null)
            {
                Order.VendID = SelectedVendor.AccNo;
            }

            // Reset inputs
            SelectedMaterial = null;
            SelectedForMaterial = null;
            PreviousRate = null;
            NewRate = 0;
            NewQty = 0;
            NewItemRemarks = string.Empty;
        }

        private void RemoveMaterial(VendOrderDetailViewModel item)
        {
            Order.OrderDetails.Remove(item);
        }

        private void EditRow(int index)
        {
            editingRowIndex = index;
        }

        private void SaveRow()
        {
            editingRowIndex = -1;
        }

        private void SelectRowMaterial(string rmid)
        {
            SelectedMaterialId = rmid;
        }
    }

    public class EmployeeLookupDto
    {
        public string EmpID { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string? DeptID { get; set; }
    }

    public class PaymentTermLookupDto
    {
        public string PaymentTerms { get; set; } = string.Empty;
    }

    public class DepartmentLookupDto
    {
        public string Name { get; set; } = string.Empty;
        public string? Description { get; set; }
    }

    public class CustomerOrderLookupDto
    {
        public string OrderNo { get; set; } = string.Empty;
    }
}
