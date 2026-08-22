using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.Production;
using Impulse.Services;
using BlazorContextMenu;

namespace Impulse.Pages.Production.MakerList
{
    public partial class MakerList
    {
        [Inject]
        public IMakerListService MakerListService { get; set; } = default!;

        [Inject]
        public NavigationManager NavigationManager { get; set; } = default!;

        [Inject]
        public Radzen.NotificationService NotificationService { get; set; } = default!;

        [Inject]
        public IReportNavigationService ReportNavigationService { get; set; } = default!;

        [Inject]
        public IBlazorContextMenuService ContextMenuService { get; set; } = default!;

        public List<MakerDto> Makers { get; set; } = new List<MakerDto>();
        public string SearchText { get; set; } = string.Empty;
        public bool IncludeInactive { get; set; } = false;
        public bool IsLoading { get; set; } = true;

        public IEnumerable<MakerDto> FilteredMakers =>
            string.IsNullOrWhiteSpace(SearchText)
                ? Makers
                : Makers.Where(m => (m.VenderName != null && m.VenderName.Contains(SearchText, StringComparison.OrdinalIgnoreCase))
                                 || (m.VendID1 != null && m.VendID1.Contains(SearchText, StringComparison.OrdinalIgnoreCase))
                                 || (m.CompanyName != null && m.CompanyName.Contains(SearchText, StringComparison.OrdinalIgnoreCase))
                                 || (m.Maker_Second_Name != null && m.Maker_Second_Name.Contains(SearchText, StringComparison.OrdinalIgnoreCase))
                                 || (m.NICNo != null && m.NICNo.Contains(SearchText, StringComparison.OrdinalIgnoreCase)));

        protected override async Task OnInitializedAsync()
        {
            await LoadMakersAsync();
        }

        private async Task LoadMakersAsync()
        {
            IsLoading = true;
            try
            {
                Makers = await MakerListService.GetMakersListAsync(IncludeInactive);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error Loading Directory",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task OnIncludeInactiveChanged(ChangeEventArgs e)
        {
            IncludeInactive = e.Value is bool b && b;
            await LoadMakersAsync();
        }

        public void NavigateToNewMaker()
        {
            NavigationManager.NavigateTo("/production/new-maker");
        }

        public void EditMaker(long vendId)
        {
            NavigationManager.NavigateTo($"/production/new-maker/{vendId}");
        }

        public async Task OpenOptionsMenu(MouseEventArgs e)
        {
            await ContextMenuService.ShowMenu("optionsBarMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        // =====================================================
        // ROW CONTEXT MENU HANDLERS
        // =====================================================
        private void OnEditRowClick(ItemClickEventArgs args)
        {
            if (args.Data is MakerDto item)
            {
                EditMaker(item.VendID);
            }
        }

        private async Task OnToggleActiveRowClick(ItemClickEventArgs args)
        {
            if (args.Data is MakerDto item)
            {
                bool newState = !item.Active;
                try
                {
                    await MakerListService.ToggleMakerActiveStatusAsync(item.VendID, newState);
                    item.Active = newState;
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Success,
                        Summary = "Status Updated",
                        Detail = $"Maker '{item.VenderName}' is now {(newState ? "Active" : "Inactive")}.",
                        Duration = 4000
                    });

                    if (!IncludeInactive && !newState)
                    {
                        await LoadMakersAsync();
                    }
                }
                catch (Exception ex)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Error,
                        Summary = "Update Failed",
                        Detail = ex.Message,
                        Duration = 5000
                    });
                }
            }
        }

        private async Task OnPrintProfileClick(ItemClickEventArgs args)
        {
            if (args.Data is MakerDto item)
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "Print_Maker_Profile.rpt",
                    SelectionFormula = $"{{VMakers.VendID}}={item.VendID}"
                });
            }
        }

        private async Task OnPrintDetailsClick(ItemClickEventArgs args)
        {
            if (args.Data is MakerDto item)
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "Print_Maker_Details.rpt",
                    SelectionFormula = $"{{VMakers.VendID}}={item.VendID}"
                });
            }
        }

        private async Task OnPrintLabelClick(ItemClickEventArgs args)
        {
            if (args.Data is MakerDto item)
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "Print_Maker_Lable.rpt",
                    SelectionFormula = $"{{VMakers.VendID}}={item.VendID}"
                });
            }
        }

        private async Task OnPrintFileLabelClick(ItemClickEventArgs args)
        {
            if (args.Data is MakerDto item)
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "Print_Maker_File_Label.rpt",
                    SelectionFormula = $"{{VMakers.VendID}}={item.VendID}"
                });
            }
        }

        private async Task OnPrintShippingLabelClick(ItemClickEventArgs args)
        {
            if (args.Data is MakerDto item)
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = "Print_Maker_Shipping_Lable.rpt",
                    SelectionFormula = $"{{VMakers.VendID}}={item.VendID}"
                });
            }
        }

        // =====================================================
        // TOP OPTIONS MENU HANDLERS
        // =====================================================
        private void OnNewMakerOptionClick(ItemClickEventArgs args)
        {
            NavigateToNewMaker();
        }

        private async Task OnRefreshListOptionClick(ItemClickEventArgs args)
        {
            await LoadMakersAsync();
            NotificationService.Notify(new Radzen.NotificationMessage
            {
                Severity = Radzen.NotificationSeverity.Info,
                Summary = "Refreshed",
                Detail = "Maker directory has been refreshed.",
                Duration = 3000
            });
        }

        private async Task OnPrintMakerListOptionClick(ItemClickEventArgs args)
        {
            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "MakerList.rpt",
                SelectionFormula = string.Empty
            });
        }
    }
}
