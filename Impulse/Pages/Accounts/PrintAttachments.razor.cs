using BlazorBootstrap;
using BlazorContextMenu;
using CIP.Configurations;
using CIP.Pages;
using CIP.Services;
using DataAccessLibrary;
using DataAccessLibrary.DAC.Accounts;
using DataAccessLibrary.Interface.Accounts;
using DataAccessLibrary.Models;
using DataAccessLibrary.Models.ViewModels.Accounts;
using DocumentFormat.OpenXml.Drawing;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Microsoft.IdentityModel.Tokens;
using Microsoft.JSInterop;
using Microsoft.VisualBasic;
using Radzen;
using System.Collections.ObjectModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Threading.Tasks;
using static System.Net.WebRequestMethods;
using OpenXmlMenu = DocumentFormat.OpenXml.Office2010.CustomUI;


namespace CIP.Pages.Accounts
{
    public partial class PrintAttachments : ComponentBase, IDisposable
    {
        private PrintAttachmentsModel statementmodel = new();

        [Inject]
        private IFinancialStatementService _financialstatementService { get; set; }
        [Inject]
        private IDBHelperService _idbHelperService { get; set; }
        [Inject]
        private IReportNavigationService ReportNavigationService { get; set; }
        [Inject]
        private INotificationService NotificationServiceManager { get; set; }
        [Inject]
        private IAuditService _auditService { get; set; }

        private List<PrintAttachmentsModel> getDBEntries { get; set; } = new List<PrintAttachmentsModel>();
        private PrintAttachmentsModel SelectedHead;
        //private PrintAttachmentsModel CurrentHead;
        private PrintAttachmentsModel CurrentHead { get; set; } = new();

        private Timer _timer;
        private bool ShowModal;
        private bool IsLoading=true;
        private bool IsAdding;
        private bool IsEdit;
        private bool bshowinactive = false;
        private bool bClearInputFile = false;
        private string myFileExtension = string.Empty;
        private IBrowserFile selectedFile;

        public string Type { get; set; } = "Expense";

        public void Dispose()
        {
            _timer?.Dispose();
        }

        private void HandleFileSelected(InputFileChangeEventArgs e)
        {
            selectedFile = e.File;
            if (selectedFile != null)
            {
                myFileExtension = System.IO.Path.GetExtension(selectedFile.Name).ToLower();
            }
        }

        private async Task UploadFileAsync(int primaryKey)
        {
            var client = ClientFactory.CreateClient("MyApiClient");
            if (selectedFile != null)
            {
                // Example: "105.pdf"
                var uniqueFileName = $"{primaryKey}{System.IO.Path.GetExtension(selectedFile.Name)}";
                uniqueFileName = uniqueFileName.Replace("/", "_");
                uniqueFileName = "Notes_" + uniqueFileName;

                using (var content = new MultipartFormDataContent())
                {
                    var fileStream = selectedFile.OpenReadStream(maxAllowedSize: 10_000_000);
                    var streamContent = new StreamContent(fileStream);

                    streamContent.Headers.ContentType =
                        new System.Net.Http.Headers.MediaTypeHeaderValue(selectedFile.ContentType);

                    content.Add(streamContent, "file", uniqueFileName);

                    var response = await client.PostAsync("api/upload", content);

                    if (!response.IsSuccessStatusCode)
                    {
                        var errorMsg = await response.Content.ReadAsStringAsync();
                        throw new InvalidOperationException($"Upload failed: {errorMsg}");
                    }
                }
            }
        }

        protected override async Task OnInitializedAsync()
        {
            IsLoading = true;
            try
            {
                await Refreshlist();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching Data: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
            }
        }
        private async Task SaveAttachment()
        {
            if (string.IsNullOrWhiteSpace(CurrentHead.Attachment_Name))
            {
                NotificationServiceManager.ShowWarning("Alert", "Cannot Save Without Attachment Name.");
                return;
            }

            if (!bClearInputFile)
            {
                if (!string.IsNullOrEmpty(myFileExtension))
                {
                    CurrentHead.Attachment_Type = myFileExtension;
                }
            }

            int entryid=await _financialstatementService.SaveNotesAttachments(CurrentHead);

            // Save item
            if (!bClearInputFile)
            {
                await UploadFileAsync(entryid);
                if (!string.IsNullOrEmpty(myFileExtension))
                {
                    CurrentHead.Attachment_Type = myFileExtension;
                }
            }

            CurrentHead.Attachment_Name = "";
            CurrentHead.Attachment_Type = "";

            await Refreshlist();
        }

        private void GoToIndexPage()
        {
            Navigation.NavigateTo("/", true);
        }

        private async Task Refreshlist()
        {
            IsLoading = true;
            try
            {
                var EntriesFromDb = await _financialstatementService.GetNotesAttachmentList();
                getDBEntries = EntriesFromDb.ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching Data: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
                StateHasChanged();
            }
        }

        [Inject]
        protected AppSettings _appSettings { get; set; }

        private void OpenAttachment(PrintAttachmentsModel item)
        {
            var fileName = "Notes_" + item.EntryID + item.Attachment_Type;
            var strUploadPath = System.IO.Path.Combine(_appSettings.UploadPath);
            var filePath = System.IO.Path.Combine(strUploadPath, $"{fileName}");

            if (!Directory.Exists(strUploadPath))
                Directory.CreateDirectory(strUploadPath);

            if (System.IO.File.Exists(filePath))
            {
                Navigation.NavigateTo($"{_appSettings.BaseAddressForReports}//cipuploads//{fileName}", true);
            }
            else
            {
                Console.WriteLine("File does not exist.");
            }

        }

    }
}