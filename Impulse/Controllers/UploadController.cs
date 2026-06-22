using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Hosting;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Extensions.Options;
using Impulse.Configurations;

namespace Impulse.Controllers
{
    [Route("api/upload")]
    [ApiController]
    public class UploadController : ControllerBase
    {
        private readonly IWebHostEnvironment _environment;
        private readonly AppSettings _appSettings;
        public UploadController(IWebHostEnvironment environment, AppSettings appSettings)
        {
            _environment = environment;
            _appSettings = appSettings;
        }

        [HttpPost]
        public async Task<IActionResult> UploadFile([FromForm] IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded.");
            }

            //var uploadPath = Path.Combine(_environment.WebRootPath, "uploads");
            //var uploadPath = @"\\sial-local\Shared Report$\CIP"; 
            var uploadPath = _appSettings.UploadPath;

            if (!Directory.Exists(uploadPath))
            {
                Directory.CreateDirectory(uploadPath);
            }

            var filePath = Path.Combine(uploadPath, file.FileName);

            using (var fileStream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(fileStream);
            }

            return Ok("File uploaded successfully.");
        }
    }
}
