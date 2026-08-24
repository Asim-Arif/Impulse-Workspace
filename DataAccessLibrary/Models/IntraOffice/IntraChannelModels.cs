using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.IntraOffice
{
    public class Channel
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public string? Description { get; set; }
        public string? DepartmentId { get; set; }
        public string? DepartmentName { get; set; }
        public bool IsPrivate { get; set; } = false;
        public bool IsActive { get; set; } = true;
        public string CreatedBy { get; set; } = string.Empty;
        public string? CreatorName { get; set; }
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        public int MemberCount { get; set; }
        public int UnreadCount { get; set; }
        public List<ChannelMember> Members { get; set; } = new List<ChannelMember>();
    }

    public class ChannelMember
    {
        public int Id { get; set; }
        public int ChannelId { get; set; }
        public string UserId { get; set; } = string.Empty;
        public string? UserName { get; set; }
        public string? FullName { get; set; }
        public string? Designation { get; set; }
        public bool IsAdmin { get; set; } = false;
        public DateTime JoinedAt { get; set; } = DateTime.UtcNow;
    }
}
