using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels.Export;

namespace Impulse.Services
{
    public class CustomInvoiceAuthService : ICustomInvoiceAuthService
    {
        private readonly ICustomInvoiceAuthDataAccess _dataAccess;
        private readonly IDBHelperService _dbHelper;
        private readonly IAuditService _auditService;

        public CustomInvoiceAuthService(ICustomInvoiceAuthDataAccess dataAccess, IDBHelperService dbHelper, IAuditService auditService)
        {
            _dataAccess = dataAccess;
            _dbHelper = dbHelper;
            _auditService = auditService;
        }

        public Task<IEnumerable<CustomerLookupModel>> GetCustomersAsync() => _dataAccess.GetCustomersAsync();

        public Task<IEnumerable<string>> GetCountriesAsync(string? customerCode) => _dataAccess.GetCountriesAsync(customerCode);

        public Task<IEnumerable<ChartOfAccountsModel>> GetSalesAccountsAsync() => _dataAccess.GetSalesAccountsAsync();

        public Task<IEnumerable<CustomInvoiceAuthItemModel>> GetPendingInvoicesAsync(string? customerCode, string? country, DateTime fromDate, DateTime toDate)
            => _dataAccess.GetPendingInvoicesAsync(customerCode, country, fromDate, toDate);

        // Legacy: strExchDiffAccNo = getGeneralDataValue("PRC_ExchDiff_AccNo")
        public async Task<string> GetExchDiffAccNoAsync() => await _dbHelper.getGeneralDataValueAsync("PRC_ExchDiff_AccNo");

        public async Task<int> AuthorizeInvoicesAsync(DateTime voucherDate, string salesAccNo, string exchDiffAccNo, List<CustomInvoiceAuthItemModel> items)
        {
            var request = new CustomInvoiceAuthSaveModel
            {
                VoucherDate = voucherDate,
                SalesAccNo = salesAccNo,
                ExchDiffAccNo = exchDiffAccNo,
                UserName = _auditService.GetCurrentUserName(),
                MachineName = _auditService.GetClientIpAddress(),
                Items = items
            };
            return await _dataAccess.AuthorizeInvoicesAsync(request);
        }
    }
}
