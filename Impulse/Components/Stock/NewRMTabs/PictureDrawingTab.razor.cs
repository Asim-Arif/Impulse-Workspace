using DataAccessLibrary.Models.ViewModels.Stock;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Forms;
using System;
using System.IO;
using System.Threading.Tasks;

namespace Impulse.Components.Stock.NewRMTabs
{
    public partial class PictureDrawingTab : ComponentBase
    {
        [Parameter]
        public NewRMViewModel Model { get; set; } = default!;

        public string PicBase64 => Model?.Pic != null ? Convert.ToBase64String(Model.Pic) : string.Empty;
        public string TechDrawingBase64 => Model?.TechnicalDrawing != null ? Convert.ToBase64String(Model.TechnicalDrawing) : string.Empty;

        // Max file size 5MB
        private const long MaxFileSize = 5 * 1024 * 1024;

        private async Task HandlePictureUpload(InputFileChangeEventArgs e)
        {
            var file = e.File;
            if (file != null)
            {
                using var stream = file.OpenReadStream(MaxFileSize);
                using var ms = new MemoryStream();
                await stream.CopyToAsync(ms);
                Model.Pic = ms.ToArray();
            }
        }

        private async Task HandleTechDrawingUpload(InputFileChangeEventArgs e)
        {
            var file = e.File;
            if (file != null)
            {
                using var stream = file.OpenReadStream(MaxFileSize);
                using var ms = new MemoryStream();
                await stream.CopyToAsync(ms);
                Model.TechnicalDrawing = ms.ToArray();
            }
        }

        private void ClearPicture()
        {
            Model.Pic = null;
        }

        private void ClearTechDrawing()
        {
            Model.TechnicalDrawing = null;
        }
    }
}
