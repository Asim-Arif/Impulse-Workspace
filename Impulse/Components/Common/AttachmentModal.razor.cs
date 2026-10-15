using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;

namespace Impulse.Components.Common
{
    public partial class AttachmentModal : ComponentBase
    {
        [Parameter] public bool IsOpen { get; set; } = false;
        [Parameter] public EventCallback<bool> IsOpenChanged { get; set; }
        [Parameter] public string EntityType { get; set; } = string.Empty;
        [Parameter] public string EntityId { get; set; } = string.Empty;
        [Parameter] public string? EntitySecondaryId { get; set; }
        [Parameter] public string? EntityDisplayName { get; set; }
        [Parameter] public EventCallback<int> OnCountChanged { get; set; }

        public void Open(string entityType, string entityId, string? secondaryId = null, string? displayName = null)
        {
            EntityType = entityType;
            EntityId = entityId;
            EntitySecondaryId = secondaryId;
            EntityDisplayName = displayName;
            IsOpen = true;
            StateHasChanged();
        }

        public async Task CloseModal()
        {
            IsOpen = false;
            if (IsOpenChanged.HasDelegate)
            {
                await IsOpenChanged.InvokeAsync(false);
            }
            StateHasChanged();
        }

        private async Task HandleCountChanged(int count)
        {
            if (OnCountChanged.HasDelegate)
            {
                await OnCountChanged.InvokeAsync(count);
            }
        }
    }
}
