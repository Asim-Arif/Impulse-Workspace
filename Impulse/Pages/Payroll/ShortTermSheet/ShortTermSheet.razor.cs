using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Constants;
using Impulse.Services;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Payroll.ShortTermSheet
{
    public partial class ShortTermSheet : ComponentBase
    {
        [Inject] private IShortTermSheetService ShortTermSheetService { get; set; } = default!;
        [Inject] private IDepartmentService DepartmentService { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;
        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; } = default!;
        [Inject] private ICommonServices CommonServices { get; set; } = default!;
        [Inject] private IVoucherService VoucherService { get; set; } = default!;
        [Inject] private IDBHelperService DBHelperService { get; set; } = default!;


        [Parameter, SupplyParameterFromQuery]
        public string? ReturnUrl { get; set; }

        private async Task ShowPrintMenu(MouseEventArgs e)
        {
            await BlazorContextMenuService.ShowMenu("printOptionsMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        private string SelectedMonthYear { get; set; } = DateTime.Today.ToString("yyyy-MM");
        private string SelectedDeptID { get; set; } = "0";
        private string SearchText { get; set; } = string.Empty;

        private List<DepartmentModel> Departments { get; set; } = new();
        private List<ShortTermSheetItemViewModel> Items { get; set; } = new();

        // ── Posting State ──────────────────────────────────────────────────
        private ShortTermPostingStatusModel PostingStatus { get; set; } = new();
        private ShortTermPostingMode PaymentMode { get; set; } = ShortTermPostingMode.Cash;

        private ChartOfAccountsModel? SelectedCashAccount { get; set; }
        private ChartOfAccountsModel? SelectedOtherAccount { get; set; }
        
        private BankAccountInfo? _selectedBankAccount;
        private BankAccountInfo? SelectedBankAccount
        {
            get => _selectedBankAccount;
            set
            {
                if (_selectedBankAccount != value)
                {
                    _selectedBankAccount = value;
                    _ = HandleBankAccountSelectionAsync(value);
                }
            }
        }

        private string ChequeNo { get; set; } = string.Empty;
        private string ChequeType { get; set; } = "Open";
        private DateTime ChequeDate { get; set; } = DateTime.Today;

        private bool DontPostVoucher { get; set; } = false;
        private bool ShowPostingSection { get; set; } = false;
        private bool IsPosting { get; set; } = false;

        private List<ChartOfAccountsModel> CashAccounts { get; set; } = new();
        private List<ChartOfAccountsModel> AllAccounts { get; set; } = new();
        private List<BankAccountInfo> BankAccounts { get; set; } = new();

        private bool IsLoading { get; set; } = false;
        private bool IsSaving { get; set; } = false;

        private DateTime SelectedDate
        {
            get
            {
                if (DateTime.TryParse(SelectedMonthYear + "-01", out DateTime dt))
                {
                    return dt;
                }
                return DateTime.Today;
            }
        }

        private DepartmentModel? _selectedDepartment = new DepartmentModel { DeptID = "0", Name = "< All Departments >" };
        private DepartmentModel? SelectedDepartment
        {
            get => _selectedDepartment;
            set
            {
                if (_selectedDepartment != value)
                {
                    _selectedDepartment = value ?? new DepartmentModel { DeptID = "0", Name = "< All Departments >" };
                    SelectedDeptID = _selectedDepartment.DeptID ?? "0";
                    _ = LoadDataAsync();
                }
            }
        }

        protected override async Task OnInitializedAsync()
        {
            _selectedDepartment = new DepartmentModel { DeptID = "0", Name = "< All Departments >" };
            await LoadDepartmentsAsync();
            await LoadAccountsAsync();
            await LoadDataAsync();
        }

        private async Task LoadDepartmentsAsync()
        {
            try
            {
                Departments = (await DepartmentService.GetDepartmentsListAsync(false)) ?? new List<DepartmentModel>();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error",
                    Detail = $"Failed to load departments: {ex.Message}"
                });
            }
        }

        private async Task LoadAccountsAsync()
        {
            try
            {
                AllAccounts = (await VoucherService.GetTransactionalAccounts()) ?? new List<ChartOfAccountsModel>();
                BankAccounts = (await VoucherService.GetBankAccounts()) ?? new List<BankAccountInfo>();

                string cashInHandParent = await DBHelperService.getSingleStringValue(
                    "DataValue", "GeneralData", "WHERE DataName='CashInHandParent'");

                if (!string.IsNullOrEmpty(cashInHandParent))
                {
                    CashAccounts = (await VoucherService.GetTransactionalAccounts($"WHERE SubAccOf='{cashInHandParent}'"))
                                   ?? new List<ChartOfAccountsModel>();
                }
                else
                {
                    CashAccounts = AllAccounts.Where(a => a.AccTitle != null && a.AccTitle.Contains("Cash", StringComparison.OrdinalIgnoreCase)).ToList();
                }

                if (!CashAccounts.Any()) CashAccounts = new List<ChartOfAccountsModel>(AllAccounts);
                if (CashAccounts.Any())
                {
                    SelectedCashAccount = CashAccounts.FirstOrDefault(a => a.AccTitle != null && a.AccTitle.Contains("Petty", StringComparison.OrdinalIgnoreCase)) ?? CashAccounts.First();
                }
            }
            catch
            {
                // Fallback for account lookups
            }
        }


        private void TogglePostingSection()
        {
            ShowPostingSection = !ShowPostingSection;
        }

        private void SetPaymentMode(ShortTermPostingMode mode)
        {
            PaymentMode = mode;
            if (mode != ShortTermPostingMode.Bank)
            {
                ChequeNo = string.Empty;
                _selectedBankAccount = null;
            }
        }

        private async Task HandleBankAccountSelectionAsync(BankAccountInfo? bank)
        {
            if (bank != null && !string.IsNullOrWhiteSpace(bank.AccNo))
            {
                try
                {
                    var chequeInfo = await VoucherService.GetNewChequeNumbersByBankAccount(bank.AccNo);
                    if (chequeInfo != null)
                    {
                        ChequeNo = chequeInfo.ChqNo ?? string.Empty;
                    }
                }
                catch { }
            }
            else
            {
                ChequeNo = string.Empty;
            }
        }


        private async Task<IEnumerable<ChartOfAccountsModel>> SearchCashAccountsAsync(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return await Task.FromResult(CashAccounts);
            return await Task.FromResult(CashAccounts.Where(a => (a.AccTitle != null && a.AccTitle.Contains(text, StringComparison.OrdinalIgnoreCase)) ||
                                                                 (a.AccNo != null && a.AccNo.Contains(text, StringComparison.OrdinalIgnoreCase))));
        }

        private async Task<IEnumerable<ChartOfAccountsModel>> SearchAllAccountsAsync(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return await Task.FromResult(AllAccounts);
            return await Task.FromResult(AllAccounts.Where(a => (a.AccTitle != null && a.AccTitle.Contains(text, StringComparison.OrdinalIgnoreCase)) ||
                                                                (a.AccNo != null && a.AccNo.Contains(text, StringComparison.OrdinalIgnoreCase))));
        }

        private async Task<IEnumerable<BankAccountInfo>> SearchBankAccountsAsync(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return await Task.FromResult(BankAccounts);
            return await Task.FromResult(BankAccounts.Where(b => (b.Bank != null && b.Bank.Contains(text, StringComparison.OrdinalIgnoreCase)) ||
                                                                 (b.AccTitle != null && b.AccTitle.Contains(text, StringComparison.OrdinalIgnoreCase))));
        }

        private async Task PostVoucherAsync()
        {
            var approvedItems = FilteredItems.Where(x => x.Approved > 0).ToList();
            if (!approvedItems.Any())
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "No employee record with approved advance amount to post."
                });
                return;
            }

            string paymentAccNo = string.Empty;
            if (PaymentMode == ShortTermPostingMode.Cash)
            {
                paymentAccNo = SelectedCashAccount?.AccNo ?? string.Empty;
            }
            else if (PaymentMode == ShortTermPostingMode.OtherAccount)
            {
                paymentAccNo = SelectedOtherAccount?.AccNo ?? string.Empty;
            }
            else if (PaymentMode == ShortTermPostingMode.Bank)
            {
                paymentAccNo = SelectedBankAccount?.AccNo ?? string.Empty;
                if (string.IsNullOrWhiteSpace(ChequeNo) && !DontPostVoucher)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Warning,
                        Summary = "Validation Warning",
                        Detail = "Please enter a valid Cheque Number for Bank Payment."
                    });
                    return;
                }
            }

            if (string.IsNullOrWhiteSpace(paymentAccNo) && !DontPostVoucher)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Please select a valid payment account."
                });
                return;
            }

            IsPosting = true;
            StateHasChanged();

            try
            {
                var dto = new ShortTermPostingDto
                {
                    Date = SelectedDate,
                    DeptID = SelectedDeptID,
                    PaymentMode = PaymentMode,
                    PaymentAccNo = paymentAccNo,
                    BankAccNo = SelectedBankAccount?.AccNo ?? string.Empty,
                    BankId = SelectedBankAccount?.AccNo ?? string.Empty,
                    ChequeNo = ChequeNo,
                    ChequeType = ChequeType,
                    ChequeDate = ChequeDate,

                    DontPostVoucher = DontPostVoucher,
                    Items = approvedItems
                };

                string vchrNo = await ShortTermSheetService.PostVouchersAndAdvancesAsync(dto);

                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Posted Successfully",
                    Detail = string.IsNullOrWhiteSpace(vchrNo)
                        ? "Advances posted to register successfully."
                        : $"Voucher {vchrNo} and Advances posted successfully.",
                    Duration = 4000
                });

                await LoadDataAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Posting Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsPosting = false;
                StateHasChanged();
            }
        }

        private async Task<IEnumerable<DepartmentModel>> SearchDepartmentsAsync(string searchText)
        {
            var allDepts = new List<DepartmentModel>
            {
                new DepartmentModel { DeptID = "0", Name = "< All Departments >" }
            };
            allDepts.AddRange(Departments);

            if (string.IsNullOrWhiteSpace(searchText))
                return await Task.FromResult(allDepts);

            return await Task.FromResult(allDepts.Where(d =>
                (d.Name != null && d.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)) ||
                (d.DeptID != null && d.DeptID.Contains(searchText, StringComparison.OrdinalIgnoreCase))));
        }


        private async Task OnMonthChanged(ChangeEventArgs e)
        {
            SelectedMonthYear = e.Value?.ToString() ?? DateTime.Today.ToString("yyyy-MM");
            await LoadDataAsync();
        }


        private async Task LoadDataAsync()
        {
            IsLoading = true;
            StateHasChanged();

            try
            {
                var date = SelectedDate;
                Items = await ShortTermSheetService.GetShortTermSheetDataAsync(date, SelectedDeptID);
                PostingStatus = await ShortTermSheetService.GetPostingStatusAsync(date, SelectedDeptID);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Data Load Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsLoading = false;
                StateHasChanged();
            }
        }

        private IEnumerable<ShortTermSheetItemViewModel> FilteredItems =>
            string.IsNullOrWhiteSpace(SearchText)
                ? Items
                : Items.Where(x => x.EmpID.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                                   x.Name.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                                   x.Designation.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                                   x.DeptName.Contains(SearchText, StringComparison.OrdinalIgnoreCase));

        private async Task SaveAsync()
        {
            if (!Items.Any()) return;

            IsSaving = true;
            StateHasChanged();

            try
            {
                var saveDtos = Items.Select(x => new ShortTermSheetSaveDto
                {
                    EmpID = x.EmpID,
                    Demand = x.Demand,
                    Approved = x.Approved
                }).ToList();

                await ShortTermSheetService.SaveShortTermSheetEntriesAsync(SelectedDate, saveDtos);

                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Success",
                    Detail = "Short Term Sheet approved amounts saved successfully.",
                    Duration = 3000
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Save Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsSaving = false;
                StateHasChanged();
            }
        }

        private async Task PrintReportAsync(bool bAllDepartments = false)
        {
            try
            {
                var dt = SelectedDate;
                var lastDayOfMonth = new DateTime(dt.Year, dt.Month, DateTime.DaysInMonth(dt.Year, dt.Month));

                string strSelection = $"{{VEmp.DT}}=#{lastDayOfMonth:yyyy-MM-dd}#";
                if (!bAllDepartments && SelectedDeptID != "0")
                {
                    strSelection += $" AND {{VEmp.DeptID}}='{SelectedDeptID}'";
                }

                var request = new ReportRequest
                {
                    ReportName = ReportNames.Payroll.AdvanceSalaryReportFilled,
                    SelectionFormula = strSelection,
                    Parameters = new Dictionary<string, object>
                    {
                        { "@DT", lastDayOfMonth },
                        { "@BankPymt", false }
                    }
                };

                await ReportNavigationService.PrintReportAsync(request);

                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Info,
                    Summary = "Generating Report",
                    Detail = "Generating Short Term Advance Sheet report...",
                    Duration = 3000
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Report Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        private async Task PrintSlipsAsync()
        {
            try
            {
                var dt = SelectedDate;
                var lastDayOfMonth = new DateTime(dt.Year, dt.Month, DateTime.DaysInMonth(dt.Year, dt.Month));

                string strSelection = $"{{VEmp.DT}}=#{lastDayOfMonth:yyyy-MM-dd}#";
                if (SelectedDeptID != "0")
                {
                    strSelection += $" AND {{VEmp.DeptID}}='{SelectedDeptID}'";
                }

                var request = new ReportRequest
                {
                    ReportName = ReportNames.Payroll.AdvanceSalarySlip,
                    SelectionFormula = strSelection
                };

                await ReportNavigationService.PrintReportAsync(request);

                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Info,
                    Summary = "Generating Slips",
                    Detail = "Generating Advance Salary Slips...",
                    Duration = 3000
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Report Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        private void GoBack()
        {
            if (!string.IsNullOrEmpty(ReturnUrl))
            {
                NavigationManager.NavigateTo(ReturnUrl);
            }
            else
            {
                NavigationManager.NavigateTo("/payroll/dashboard");
            }
        }
    }
}
