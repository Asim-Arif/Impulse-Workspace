using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Radzen;

namespace Impulse.Components.Payroll.NewEmployeeTabs
{
    public partial class TargetsTab
    {
        [Parameter]
        public EmployeeDto Employee { get; set; } = default!;

        [Parameter]
        public List<ProcessLookupModel> Processes { get; set; } = new();

        [Parameter]
        public List<GroupLookupModel> Groups { get; set; } = new();

        [Inject]
        public NotificationService NotificationService { get; set; } = default!;

        protected int _newProcessID;
        protected int _newGroupID;
        protected int? _newQty;
        protected bool _newOverTime;

        protected ProcessLookupModel? _dummyProcessForGrid = new();
        protected GroupLookupModel? _dummyGroupForGrid = new();

        protected bool _isCopyModalOpen = false;

        protected ProcessLookupModel? SelectedNewProcess => Processes.FirstOrDefault(p => p.ProcessID == _newProcessID);
        protected GroupLookupModel? SelectedNewGroup => Groups.FirstOrDefault(g => g.GroupID == _newGroupID);

        protected void OnNewProcessChanged(ProcessLookupModel? val)
        {
            _newProcessID = val?.ProcessID ?? 0;
        }

        protected void OnNewGroupChanged(GroupLookupModel? val)
        {
            _newGroupID = val?.GroupID ?? 0;
        }

        protected Task<IEnumerable<ProcessLookupModel>> SearchProcesses(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<ProcessLookupModel>>(Processes);
            return Task.FromResult<IEnumerable<ProcessLookupModel>>(
                Processes.Where(p => p.Description.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList()
            );
        }

        protected Task<IEnumerable<GroupLookupModel>> SearchGroups(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<GroupLookupModel>>(Groups);
            return Task.FromResult<IEnumerable<GroupLookupModel>>(
                Groups.Where(g => g.Description.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList()
            );
        }

        protected void AddTarget()
        {
            if (_newProcessID <= 0)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Please select a Process.",
                    Duration = 3000
                });
                return;
            }

            if (_newGroupID <= 0)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Validation",
                    Detail = "Please select a Group.",
                    Duration = 3000
                });
                return;
            }

            Employee.DailyTargets.Add(new EmpDailyTargetDto
            {
                ProcessID = _newProcessID,
                GroupID = _newGroupID,
                Qty = _newQty,
                OverTime = _newOverTime ? (byte)1 : (byte)0
            });

            _newProcessID = 0;
            _newGroupID = 0;
            _newQty = null;
            _newOverTime = false;
        }

        [Parameter]
        public EventCallback OnOpenCopyModal { get; set; }

        protected async Task OpenCopyModal()
        {
            if (OnOpenCopyModal.HasDelegate)
            {
                await OnOpenCopyModal.InvokeAsync();
            }
            else
            {
                _isCopyModalOpen = true;
            }
        }

        protected void HandleTargetsCopied((List<EmpDailyTargetDto> Targets, bool ReplaceExisting) args)
        {
            if (args.Targets == null || args.Targets.Count == 0) return;

            if (args.ReplaceExisting)
            {
                Employee.DailyTargets.Clear();
                Employee.DailyTargets.AddRange(args.Targets);
            }
            else
            {
                int addedCount = 0;
                foreach (var t in args.Targets)
                {
                    // Avoid duplicating exact same Process and Group combination
                    bool exists = Employee.DailyTargets.Any(x => x.ProcessID == t.ProcessID && x.GroupID == t.GroupID && x.OverTime == t.OverTime);
                    if (!exists)
                    {
                        Employee.DailyTargets.Add(t);
                        addedCount++;
                    }
                }
            }

            NotificationService.Notify(new NotificationMessage
            {
                Severity = NotificationSeverity.Success,
                Summary = "Targets Copied",
                Detail = args.ReplaceExisting
                    ? $"Replaced current targets with {args.Targets.Count} copied target(s)."
                    : $"Appended targets to list successfully.",
                Duration = 4000
            });

            StateHasChanged();
        }
    }
}
