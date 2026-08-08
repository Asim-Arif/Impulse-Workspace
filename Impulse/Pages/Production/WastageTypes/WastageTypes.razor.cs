using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.Production.WastageTypes
{
    public partial class WastageTypes : ComponentBase
    {
        [Inject] public IWastageTypeService WastageTypeService { get; set; } = default!;
        [Inject] public Radzen.NotificationService NotificationService { get; set; } = default!;

        public List<WastageTypeModel> WastageTypesList { get; set; } = new();
        public WastageTypeModel? SelectedWastageType { get; set; }
        public WastageTypeModel? ContextRowItem { get; set; }

        public string SearchText { get; set; } = string.Empty;
        public bool IsLoading { get; set; } = false;

        // Modal States
        public bool ShowModal { get; set; } = false;
        public WastageTypeDto? ModelForEdit { get; set; }

        // Delete Modal
        public bool ShowDeleteModal { get; set; } = false;
        public WastageTypeModel? ItemToDelete { get; set; }
        public bool IsDeleting { get; set; } = false;

        public List<WastageTypeModel> FilteredWastageTypes
        {
            get
            {
                if (string.IsNullOrWhiteSpace(SearchText))
                    return WastageTypesList;

                string q = SearchText.Trim();
                return WastageTypesList.Where(w =>
                    (!string.IsNullOrEmpty(w.WastageName) && w.WastageName.Contains(q, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(w.WastageTypeTitle) && w.WastageTypeTitle.Contains(q, StringComparison.OrdinalIgnoreCase))
                ).ToList();
            }
        }

        protected override async Task OnInitializedAsync()
        {
            await LoadWastageTypesAsync();
        }

        public async Task LoadWastageTypesAsync()
        {
            IsLoading = true;
            try
            {
                WastageTypesList = await WastageTypeService.GetWastageTypesAsync();
                if (SelectedWastageType != null)
                {
                    SelectedWastageType = WastageTypesList.FirstOrDefault(w => w.EntryID == SelectedWastageType.EntryID);
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Failed to Load Wastage Types",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        public void OpenCreateModal()
        {
            ModelForEdit = new WastageTypeDto
            {
                WastageType = 0,
                Closed = false
            };
            ShowModal = true;
        }

        public async Task OpenEditModal(WastageTypeModel item)
        {
            if (item == null) return;
            try
            {
                var dto = await WastageTypeService.GetWastageTypeByIdAsync(item.EntryID);
                if (dto != null)
                {
                    ModelForEdit = dto;
                    ShowModal = true;
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Wastage Type",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public async Task OnModalSaved()
        {
            ShowModal = false;
            ModelForEdit = null;
            await LoadWastageTypesAsync();
        }

        public void OnModalClosed()
        {
            ShowModal = false;
            ModelForEdit = null;
        }

        public async Task PromptDeleteWastageType(WastageTypeModel item)
        {
            if (item == null) return;

            var (canDelete, reason) = await WastageTypeService.CanDeleteWastageTypeAsync(item.EntryID);
            if (!canDelete)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Cannot Delete Wastage Type",
                    Detail = reason,
                    Duration = 5000
                });
                return;
            }

            ItemToDelete = item;
            ShowDeleteModal = true;
        }

        public async Task ConfirmDeleteWastageType()
        {
            if (ItemToDelete == null) return;

            IsDeleting = true;
            try
            {
                bool ok = await WastageTypeService.DeleteWastageTypeAsync(ItemToDelete.EntryID);
                if (ok)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Wastage Type Deleted",
                        Detail = $"Wastage Type '{ItemToDelete.WastageName}' deleted successfully.",
                        Duration = 3000
                    });

                    ShowDeleteModal = false;
                    ItemToDelete = null;
                    await LoadWastageTypesAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Delete Failed",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsDeleting = false;
            }
        }

        // Context Menu Handlers
        public void ContextNew(ItemClickEventArgs args) => OpenCreateModal();

        public async Task ContextEdit(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                await OpenEditModal(ContextRowItem);
            }
        }

        public async Task ContextDelete(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                await PromptDeleteWastageType(ContextRowItem);
            }
        }

        private void ResolveRowItem(ItemClickEventArgs args)
        {
            if (args.Data is WastageTypeModel item)
            {
                ContextRowItem = item;
                SelectedWastageType = item;
            }
        }
    }
}
