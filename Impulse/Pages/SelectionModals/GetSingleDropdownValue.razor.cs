using DataAccessLibrary.Models.ViewModels;
using Microsoft.AspNetCore.Components;

namespace Impulse.Pages.SelectionModals
{
    public partial class GetSingleDropdownValue
    {
        [Parameter] public string Title { get; set; } = "Select Value";
        [Parameter] public string Label { get; set; } = "Choose an option";
        [Parameter] public List<GenericDropDownModel> Items { get; set; } = new();
        [Parameter] public EventCallback<string> OnConfirmed { get; set; }

        protected bool IsVisible { get; set; }
        protected string SelectedValue { get; set; }

        public void Show()
        {
            IsVisible = true;
            SelectedValue = "";
            
            StateHasChanged();
        }

        public void Hide()
        {
            IsVisible = false;
            StateHasChanged();
        }

        protected void Cancel() => Hide();

        protected async Task Confirm()
        {
            await OnConfirmed.InvokeAsync(SelectedValue);
            Hide();
        }
    }
}