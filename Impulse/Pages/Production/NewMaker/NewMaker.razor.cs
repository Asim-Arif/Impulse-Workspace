using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.Production;

namespace Impulse.Pages.Production.NewMaker
{
    public partial class NewMaker
    {
        [Parameter]
        public long? VendId { get; set; }

        [Inject]
        public IMakerService MakerService { get; set; } = default!;

        [Inject]
        public NavigationManager NavigationManager { get; set; } = default!;

        [Inject]
        public Radzen.NotificationService NotificationService { get; set; } = default!;

        public MakerDto Model { get; set; } = new MakerDto();
        public List<MakerCategoryModel> Categories { get; set; } = new List<MakerCategoryModel>();
        public MakerCategoryModel? SelectedCategory { get; set; }
        public List<string> PaymentTermsList { get; set; } = new List<string>();
        public List<ProcessModel> Processes { get; set; } = new List<ProcessModel>();

        public string ActiveTab { get; set; } = "GeneralInfo";
        public bool IsSaving { get; set; } = false;
        public bool IsEditMode => VendId.HasValue && VendId.Value > 0;

        protected override async Task OnInitializedAsync()
        {
            await LoadLookupsAsync();

            if (IsEditMode)
            {
                var existing = await MakerService.GetMakerByIdAsync(VendId!.Value);
                if (existing != null)
                {
                    Model = existing;
                    SelectedCategory = Categories.FirstOrDefault(c => c.AccNo == Model.SubAccOf);
                    var assignedIds = await MakerService.GetMakerProcessIdsAsync(VendId.Value);
                    var assignedSet = new HashSet<int>(assignedIds);

                    foreach (var proc in Processes)
                    {
                        proc.IsSelected = assignedSet.Contains(proc.ProcessID);
                    }
                }
                else
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Error,
                        Summary = "Error",
                        Detail = "Maker record not found.",
                        Duration = 4000
                    });
                    NavigationManager.NavigateTo("/production");
                }
            }
            else
            {
                Model.VendID1 = await MakerService.GenerateNextMakerCodeAsync();
            }
        }

        private async Task LoadLookupsAsync()
        {
            Categories = await MakerService.GetMakerCategoriesAsync();
            PaymentTermsList = await MakerService.GetPaymentTermsAsync();
            Processes = await MakerService.GetAllProcessesAsync();
        }

        public async Task OnCategorySelected(MakerCategoryModel cat)
        {
            SelectedCategory = cat;
            if (cat != null)
            {
                Model.SubAccOf = cat.AccNo;
                if (!IsEditMode)
                {
                    Model.AccNo = await MakerService.GenerateNextAccNoAsync(cat.AccNo);
                }
            }
            else
            {
                Model.SubAccOf = null;
                if (!IsEditMode)
                {
                    Model.AccNo = string.Empty;
                }
            }
        }

        public Task<IEnumerable<MakerCategoryModel>> SearchCategories(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<MakerCategoryModel>>(Categories);

            return Task.FromResult<IEnumerable<MakerCategoryModel>>(
                Categories.Where(c => c.AccTitle.Contains(searchText, StringComparison.OrdinalIgnoreCase)
                                   || c.AccNo.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        public void ChangeTab(string tab)
        {
            ActiveTab = tab;
        }

        public async Task HandleValidSubmit()
        {
            await Save(closeAfter: false);
        }

        public async Task Save(bool closeAfter)
        {
            if (string.IsNullOrWhiteSpace(Model.VenderName))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Please enter Maker / Company Name.",
                    Duration = 4000
                });
                return;
            }

            if (string.IsNullOrWhiteSpace(Model.NICNo))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Please enter NIC Number.",
                    Duration = 4000
                });
                return;
            }

            if (SelectedCategory == null || string.IsNullOrWhiteSpace(Model.SubAccOf))
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Warning,
                    Summary = "Validation Warning",
                    Detail = "Please select Maker Category.",
                    Duration = 4000
                });
                return;
            }

            // Check NIC Duplicate
            bool isDuplicate = await MakerService.IsNicDuplicateAsync(Model.NICNo, VendId);
            if (isDuplicate)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Duplicate NIC",
                    Detail = "A maker with this NIC Number already exists in the system.",
                    Duration = 5000
                });
                return;
            }

            IsSaving = true;
            try
            {
                var selectedProcessIds = Processes.Where(p => p.IsSelected).Select(p => p.ProcessID).ToList();

                if (IsEditMode)
                {
                    await MakerService.UpdateMakerAsync(Model, selectedProcessIds);
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Success,
                        Summary = "Success",
                        Detail = "Maker information updated successfully.",
                        Duration = 4000
                    });
                }
                else
                {
                    long newVendId = await MakerService.SaveMakerAsync(Model, selectedProcessIds);
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Success,
                        Summary = "Success",
                        Detail = $"Maker '{Model.VenderName}' created successfully with Code {Model.VendID1}.",
                        Duration = 4000
                    });
                }

                if (closeAfter)
                {
                    NavigationManager.NavigateTo("/production");
                }
                else
                {
                    // Reset form for next entry
                    Model = new MakerDto();
                    SelectedCategory = null;
                    foreach (var p in Processes) p.IsSelected = false;
                    Model.VendID1 = await MakerService.GenerateNextMakerCodeAsync();
                    ActiveTab = "GeneralInfo";
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Save Failed",
                    Detail = ex.Message,
                    Duration = 6000
                });
            }
            finally
            {
                IsSaving = false;
            }
        }

        public void Cancel()
        {
            NavigationManager.NavigateTo("/production");
        }
    }
}
