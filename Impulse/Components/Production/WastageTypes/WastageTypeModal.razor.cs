using System;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Components.Production.WastageTypes
{
    public partial class WastageTypeModal : ComponentBase
    {
        [Inject] public IWastageTypeService WastageTypeService { get; set; } = default!;
        [Inject] public Radzen.NotificationService NotificationService { get; set; } = default!;

        [Parameter] public bool IsOpen { get; set; }
        [Parameter] public WastageTypeDto? Model { get; set; }
        [Parameter] public EventCallback OnSaved { get; set; }
        [Parameter] public EventCallback OnClosed { get; set; }

        public bool IsSaving { get; set; } = false;

        public async Task SaveWastageTypeAsync()
        {
            if (Model == null) return;

            if (string.IsNullOrWhiteSpace(Model.WastageName))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Wastage Name Required",
                    Detail = "Please provide a valid wastage type name.",
                    Duration = 3000
                });
                return;
            }

            IsSaving = true;
            try
            {
                if (Model.EntryID == 0)
                {
                    int newId = await WastageTypeService.SaveWastageTypeAsync(Model);
                    if (newId > 0)
                    {
                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Success,
                            Summary = "Wastage Type Created",
                            Detail = $"Wastage Type '{Model.WastageName}' created successfully.",
                            Duration = 4000
                        });

                        IsOpen = false;
                        await OnSaved.InvokeAsync();
                    }
                }
                else
                {
                    bool ok = await WastageTypeService.UpdateWastageTypeAsync(Model);
                    if (ok)
                    {
                        NotificationService.Notify(new NotificationMessage
                        {
                            Severity = NotificationSeverity.Success,
                            Summary = "Wastage Type Updated",
                            Detail = $"Wastage Type '{Model.WastageName}' updated successfully.",
                            Duration = 4000
                        });

                        IsOpen = false;
                        await OnSaved.InvokeAsync();
                    }
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Save Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsSaving = false;
            }
        }

        public async Task CloseModal()
        {
            IsOpen = false;
            await OnClosed.InvokeAsync();
        }
    }
}
