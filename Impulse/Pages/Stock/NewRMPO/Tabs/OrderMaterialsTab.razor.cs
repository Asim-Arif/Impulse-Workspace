using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Stock;
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
        private NotificationService NotificationService { get; set; }

        private List<NewVendorViewModel> Vendors = new();
        private List<RawMaterialLookupModel> AllMaterials = new();

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
            }
        }
        public RawMaterialLookupModel SelectedMaterial { get; set; }
        public RawMaterialLookupModel SelectedForMaterial { get; set; }
        public bool ShowCCItems { get; set; }
        public float NewRate { get; set; }
        public float NewQty { get; set; }
        public string NewItemRemarks { get; set; }

        private int editingRowIndex = -1;

        protected override async Task OnInitializedAsync()
        {
            var vendorsResult = await VendorService.GetVendorsAsync();
            Vendors = vendorsResult != null ? vendorsResult : new List<NewVendorViewModel>();
            
            var materialsResult = await RMService.GetAvailableRawMaterialsAsync();
            AllMaterials = materialsResult != null ? materialsResult : new List<RawMaterialLookupModel>();

            if (Order != null && !string.IsNullOrEmpty(Order.VendID))
            {
                SelectedVendor = Vendors.FirstOrDefault(v => v.AccNo == Order.VendID);
            }
        }

        private async Task<IEnumerable<NewVendorViewModel>> SearchVendors(string searchText)
        {
            await Task.Delay(50);
            if (string.IsNullOrEmpty(searchText))
                return Vendors;

            return Vendors.Where(v => v.VendorName.Contains(searchText, System.StringComparison.OrdinalIgnoreCase)).ToList();
        }

        private async Task<IEnumerable<RawMaterialLookupModel>> SearchMaterials(string searchText)
        {
            await Task.Delay(50);
            var query = AllMaterials.AsEnumerable();

            if (ShowCCItems)
            {
                // In a real scenario, there would be a flag like IsCCItem or we filter by Sample=1 based on legacy code
                // Assuming RawMaterialLookupModel has a way to identify this, else we just return all for now
            }

            if (Order.POType == 3) // Male Grinding
            {
                // Filter specifically if needed
            }

            if (string.IsNullOrEmpty(searchText))
                return query.Take(100).ToList();

            return query.Where(m => m.RMName.Contains(searchText, System.StringComparison.OrdinalIgnoreCase)).Take(100).ToList();
        }

        private void OnPOTypeChanged()
        {
            // Reset selected items if needed
            SelectedMaterial = null;
            SelectedForMaterial = null;
        }

        private void OnShowCCItemsChanged()
        {
            SelectedMaterial = null;
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
                ForRMID = SelectedForMaterial != null && int.TryParse(SelectedForMaterial.RMID1, out int forRmId) ? forRmId : (int?)null
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
    }
}
