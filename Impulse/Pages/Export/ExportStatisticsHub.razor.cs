using System.Collections.Generic;
using Microsoft.AspNetCore.Components;
using Impulse.Services;

namespace Impulse.Pages.Export
{
    public partial class ExportStatisticsHub : ComponentBase
    {
        [Inject] private INotificationService Notification { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;

        public class StatisticItem
        {
            public string Name { get; set; } = string.Empty;
            public string Icon { get; set; } = string.Empty;
        }

        private readonly List<StatisticItem> statistics = new()
        {
            new StatisticItem { Name = "Total Export", Icon = "fas fa-ship" },
            new StatisticItem { Name = "Total Demand", Icon = "fas fa-chart-line" },
            new StatisticItem { Name = "Articlewise Sales", Icon = "fas fa-chart-pie" },
            new StatisticItem { Name = "Articlewise Demand", Icon = "fas fa-chart-bar" },
            new StatisticItem { Name = "Itemwise Export Performance", Icon = "fas fa-tachometer-alt" },
            new StatisticItem { Name = "Customer Demand Report Current Year (Qty)", Icon = "fas fa-boxes" },
            new StatisticItem { Name = "Customer Demand Report Current Year (Amt)", Icon = "fas fa-dollar-sign" },
            new StatisticItem { Name = "Current Year monthwise demand", Icon = "fas fa-calendar-alt" },
            new StatisticItem { Name = "Yearly Demand/Supply", Icon = "fas fa-history" },
            new StatisticItem { Name = "Demand/Supply Comparison Graph Qty wise", Icon = "fas fa-chart-area" },
            new StatisticItem { Name = "Demand/Supply Comparison Graph Amt wise", Icon = "fas fa-chart-line" },
            new StatisticItem { Name = "Main Group Demand/Supply", Icon = "fas fa-folder" },
            new StatisticItem { Name = "Sub-Group Demand/Supply", Icon = "fas fa-folder-open" },
            new StatisticItem { Name = "Main Group/Sub-Group Item Level D&&S", Icon = "fas fa-indent" },
            new StatisticItem { Name = "Customerwise Demands", Icon = "fas fa-users" },
            new StatisticItem { Name = "Total Export Value Share", Icon = "fas fa-percentage" },
            new StatisticItem { Name = "Quarterly Demand", Icon = "fas fa-calendar-check" },
            new StatisticItem { Name = "Delivery Performance Report", Icon = "fas fa-truck" },
            new StatisticItem { Name = "Delivery Performance Detail", Icon = "fas fa-info-circle" },
            new StatisticItem { Name = "Customer Purchase History", Icon = "fas fa-history" },
            new StatisticItem { Name = "Itemwise Yearly Demand", Icon = "fas fa-list-ol" },
            new StatisticItem { Name = "Itemwise Yearly Demand (X2)", Icon = "fas fa-list-ul" },
            new StatisticItem { Name = "Itemwise Yearwise Demand", Icon = "fas fa-th-list" },
            new StatisticItem { Name = "Yearly Customerwise Demands", Icon = "fas fa-user-friends" },
            new StatisticItem { Name = "Product Line Demand", Icon = "fas fa-stream" },
            new StatisticItem { Name = "Product Groupwise Demand", Icon = "fas fa-layer-group" },
            new StatisticItem { Name = "Product Line Supply", Icon = "fas fa-truck-loading" },
            new StatisticItem { Name = "Product Groupwise Supply", Icon = "fas fa-network-wired" },
            new StatisticItem { Name = "Total Demand Value Share", Icon = "fas fa-chart-pie" },
            new StatisticItem { Name = "Advance Plan", Icon = "fas fa-calendar-plus" }
        };

        private void OnCardClick(string name)
        {
            if (name == "Total Export")
            {
                NavigationManager.NavigateTo("/export/statistics/total-export");
            }
            else if (name == "Total Demand")
            {
                NavigationManager.NavigateTo("/export/statistics/total-demand");
            }
            else if (name == "Articlewise Sales")
            {
                NavigationManager.NavigateTo("/export/statistics/articlewise-sales");
            }
            else if (name == "Articlewise Demand")
            {
                NavigationManager.NavigateTo("/export/statistics/articlewise-demand");
            }
            else
            {
                Notification.ShowInformation("Statistics", $"The '{name}' report has not been migrated yet.");
            }
        }
    }
}
