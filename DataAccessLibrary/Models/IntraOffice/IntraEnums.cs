namespace DataAccessLibrary.Models.IntraOffice
{
    public enum AnnouncementPriority
    {
        Normal = 0,
        Important = 1,
        Critical = 2
    }

    public enum MeetingStatus
    {
        Scheduled = 0,
        InProgress = 1,
        Completed = 2,
        Cancelled = 3
    }

    public enum MessageType
    {
        Text = 0,
        File = 1,
        Audio = 2,
        System = 3
    }

    public enum TaskItemStatus
    {
        Pending = 0,
        InProgress = 1,
        Completed = 2,
        Cancelled = 3
    }

    public enum TaskPriority
    {
        Low = 0,
        Normal = 1,
        High = 2,
        Urgent = 3
    }

    public enum UserStatus
    {
        Offline = 0,
        Online = 1,
        Away = 2,
        Busy = 3
    }

    public enum ApproverType
    {
        Manager = 0,
        Role = 1,
        SpecificUser = 2
    }
}
