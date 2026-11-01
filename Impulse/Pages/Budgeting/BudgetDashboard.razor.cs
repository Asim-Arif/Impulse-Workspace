using System;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Budgeting;
using DataAccessLibrary.Models.ViewModels.Budgeting;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.Budgeting
{
    public partial class BudgetDashboard : ComponentBase
    {
        [Inject] public IBudgetDashboardDataAccess DataAccess { get; set; } = default!;
        [Inject] public NavigationManager Navigation { get; set; } = default!;
        [Inject] public NotificationService NotificationService { get; set; } = default!;

        protected BudgetDashboardFilter Filter { get; set; } = new();
        protected BudgetDashboardResult Data { get; set; } = new();
        protected bool IsLoading { get; set; } = true;

        protected static readonly (int Index, string Name, string ShortName)[] MonthList = new[]
        {
            (0, "All Months", "All"),
            (1, "July", "Jul"),
            (2, "August", "Aug"),
            (3, "September", "Sep"),
            (4, "October", "Oct"),
            (5, "November", "Nov"),
            (6, "December", "Dec"),
            (7, "January", "Jan"),
            (8, "February", "Feb"),
            (9, "March", "Mar"),
            (10, "April", "Apr"),
            (11, "May", "May"),
            (12, "June", "Jun")
        };

        protected static readonly (int Index, string Name, string ShortName)[] QuarterList = new[]
        {
            (0, "Full Year", "FY"),
            (1, "Q1: Jul - Sep", "Q1"),
            (2, "Q2: Oct - Dec", "Q2"),
            (3, "Q3: Jan - Mar", "Q3"),
            (4, "Q4: Apr - Jun", "Q4")
        };

        protected override async Task OnInitializedAsync()
        {
            await LoadDashboardDataAsync();
        }

        protected async Task LoadDashboardDataAsync()
        {
            try
            {
                IsLoading = true;
                Data = await DataAccess.GetDashboardDataAsync(Filter);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Dashboard Load Error",
                    Detail = "Failed to load budgeting data: " + ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        protected async Task SelectYearAsync(string fy)
        {
            Filter.FinancialYear = fy;
            await LoadDashboardDataAsync();
        }

        protected async Task SelectMonthAsync(int monthIndex)
        {
            Filter.MonthIndex = monthIndex;
            Filter.QuarterIndex = 0;
            await LoadDashboardDataAsync();
        }

        protected async Task SelectQuarterAsync(int quarterIndex)
        {
            Filter.QuarterIndex = quarterIndex;
            Filter.MonthIndex = 0;
            await LoadDashboardDataAsync();
        }

        protected async Task SetModeAsync(string mode)
        {
            Filter.Mode = mode;
            Filter.HeadCode = "0";
            await LoadDashboardDataAsync();
        }

        protected string SelectedPeriodLabel =>
            Filter.MonthIndex >= 1 && Filter.MonthIndex < MonthList.Length
                ? MonthList[Filter.MonthIndex].Name
                : QuarterList.FirstOrDefault(x => x.Index == Filter.QuarterIndex).Name ?? "Full Year";

        protected static string FormatCurrency(decimal val)
        {
            return val.ToString("N0");
        }

        protected static string FormatShortNumber(decimal val)
        {
            if (Math.Abs(val) >= 1_000_000_000)
                return (val / 1_000_000_000m).ToString("0.##") + " B";
            if (Math.Abs(val) >= 1_000_000)
                return (val / 1_000_000m).ToString("0.##") + " M";
            if (Math.Abs(val) >= 1_000)
                return (val / 1_000m).ToString("0.##") + " K";
            return val.ToString("N0");
        }
    }
}

