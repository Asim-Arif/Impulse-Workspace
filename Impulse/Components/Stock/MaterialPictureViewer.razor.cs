using Microsoft.AspNetCore.Components;
using System;
using System.Threading.Tasks;

namespace Impulse.Components.Stock
{
    public partial class MaterialPictureViewer : ComponentBase
    {
        [Parameter] public string MaterialId { get; set; }
        [Parameter] public bool ShowPicture { get; set; }

        private string ImageBase64;
        private bool IsLoading = false;
        private string LastLoadedMaterialId;

        protected override async Task OnParametersSetAsync()
        {
            if (ShowPicture && !string.IsNullOrEmpty(MaterialId) && MaterialId != LastLoadedMaterialId)
            {
                LastLoadedMaterialId = MaterialId;
                await LoadPictureAsync();
            }
            else if (!ShowPicture || string.IsNullOrEmpty(MaterialId))
            {
                ImageBase64 = null;
                LastLoadedMaterialId = null;
            }
        }

        private async Task LoadPictureAsync()
        {
            IsLoading = true;
            StateHasChanged();
            
            try
            {
                byte[] imageBytes = await RMService.GetMaterialPictureAsync(MaterialId);
                if (imageBytes != null && imageBytes.Length > 0)
                {
                    ImageBase64 = Convert.ToBase64String(imageBytes);
                }
                else
                {
                    ImageBase64 = null;
                }
            }
            catch
            {
                ImageBase64 = null;
            }
            finally
            {
                IsLoading = false;
            }
        }
    }
}
