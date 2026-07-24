using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.AspNetCore.Components;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Components.Stock.NewRMTabs
{
    public partial class AccountInfoTab : ComponentBase
    {
        [Parameter]
        public NewRMViewModel Model { get; set; } = default!;

        // These would typically be passed in or loaded from a service in a real app,
        // but to keep the tab self-contained, we accept them as parameters from the parent.
        [Parameter]
        public List<ChartOfAccountsModel> SaleAccounts { get; set; } = new();

        [Parameter]
        public List<ChartOfAccountsModel> SaleRAccounts { get; set; } = new();

        [Parameter]
        public List<ChartOfAccountsModel> AssetAccounts { get; set; } = new();

        public ChartOfAccountsModel? SelectedSaleAcc { get; set; }
        public ChartOfAccountsModel? SelectedSaleRAcc { get; set; }
        public ChartOfAccountsModel? SelectedAssetAcc { get; set; }

        protected override void OnParametersSet()
        {
            if (Model != null)
            {
                if (!string.IsNullOrWhiteSpace(Model.SaleAccNo) && SaleAccounts != null)
                {
                    SelectedSaleAcc = SaleAccounts.FirstOrDefault(x => x.AccNo == Model.SaleAccNo);
                }

                if (!string.IsNullOrWhiteSpace(Model.SRTAccNo) && SaleRAccounts != null)
                {
                    SelectedSaleRAcc = SaleRAccounts.FirstOrDefault(x => x.AccNo == Model.SRTAccNo);
                }

                if (!string.IsNullOrWhiteSpace(Model.AssetAccNo) && AssetAccounts != null)
                {
                    SelectedAssetAcc = AssetAccounts.FirstOrDefault(x => x.AccNo == Model.AssetAccNo);
                }
            }
        }

        private void OnSaleAccChanged(ChartOfAccountsModel? acc)
        {
            SelectedSaleAcc = acc;
            Model.SaleAccNo = acc?.AccNo ?? "";
        }

        private void OnSaleRAccChanged(ChartOfAccountsModel? acc)
        {
            SelectedSaleRAcc = acc;
            Model.SRTAccNo = acc?.AccNo ?? "";
        }

        private void OnAssetAccChanged(ChartOfAccountsModel? acc)
        {
            SelectedAssetAcc = acc;
            Model.AssetAccNo = acc?.AccNo ?? "";
        }

        private Task<IEnumerable<ChartOfAccountsModel>> SearchSaleAccounts(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult(SaleAccounts.Take(50));

            return Task.FromResult(SaleAccounts
                .Where(x => 
                    (x.AccNo?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false) ||
                    (x.AccTitle?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false))
                .Take(50));
        }

        private Task<IEnumerable<ChartOfAccountsModel>> SearchSaleRAccounts(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult(SaleRAccounts.Take(50));

            return Task.FromResult(SaleRAccounts
                .Where(x => 
                    (x.AccNo?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false) ||
                    (x.AccTitle?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false))
                .Take(50));
        }

        private Task<IEnumerable<ChartOfAccountsModel>> SearchAssetAccounts(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult(AssetAccounts.Take(50));

            return Task.FromResult(AssetAccounts
                .Where(x => 
                    (x.AccNo?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false) ||
                    (x.AccTitle?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false))
                .Take(50));
        }
    }
}
