using System;
using System.Threading.Tasks;
using Impulse.Services.Setup;
using Microsoft.AspNetCore.Components;

namespace Impulse.Pages.Setups
{
    public partial class SetupDashboard : ComponentBase
    {
        [Inject]
        protected IUserService UserService { get; set; } = default!;

        protected int TotalUsers { get; set; } = 0;
        protected int ActiveUsers { get; set; } = 0;
        protected int InactiveUsers { get; set; } = 0;
        protected bool IsLoading { get; set; } = true;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                IsLoading = true;
                var stats = await UserService.GetUserStatisticsAsync();
                TotalUsers = stats.TotalUsers;
                ActiveUsers = stats.ActiveUsers;
                InactiveUsers = stats.InactiveUsers;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading setup dashboard stats: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
            }
        }
    }
}
