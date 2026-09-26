using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class HubSupervisorInfoDto
    {
        public int UserID { get; set; }
        public string UserName { get; set; } = string.Empty;
        public string? EmpID { get; set; }
        public string? EmployeeName { get; set; }
        public string? Designation { get; set; }
    }

    public class HubDetectionResultDto
    {
        public int GroupID { get; set; }
        public string GroupName { get; set; } = string.Empty;

        public int CurrentProcessId { get; set; }
        public string CurrentProcessName { get; set; } = string.Empty;
        public int CurrentSeqNo { get; set; }
        public string CurrentHubName { get; set; } = string.Empty;

        public bool IsLastProcessOfHub { get; set; }
        public bool IsFinalGroupHub { get; set; }

        public string? NextHubName { get; set; }
        public int? NextFirstProcessId { get; set; }
        public string? NextFirstProcessName { get; set; }

        public List<HubSupervisorInfoDto> NextHubSupervisors { get; set; } = new();
    }

    public class HubTransitionExecutionResult
    {
        public bool Handled { get; set; }
        public bool IsLastProcessOfHub { get; set; }
        public bool IsFinalHub { get; set; }
        public string CurrentHubName { get; set; } = string.Empty;
        public string? NextHubName { get; set; }
        public int TasksClosed { get; set; }
        public int TasksCreated { get; set; }
        public int NotificationsSent { get; set; }
        public string Message { get; set; } = string.Empty;
    }
}
