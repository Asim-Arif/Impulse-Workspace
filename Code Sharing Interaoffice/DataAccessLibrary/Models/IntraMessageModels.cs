using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.IntraOffice
{
    public class Message
    {
        public long Id { get; set; }
        public int? ChannelId { get; set; }
        public string SenderId { get; set; } = string.Empty;
        public string? SenderName { get; set; }
        public string? SenderDesignation { get; set; }
        public string? ReceiverId { get; set; }
        public string? ReceiverName { get; set; }
        public string Content { get; set; } = string.Empty;
        public MessageType MessageType { get; set; } = MessageType.Text;
        public bool IsRead { get; set; } = false;
        public bool IsDeleted { get; set; } = false;
        public long? ParentMessageId { get; set; }
        public string? ParentMessagePreview { get; set; }
        public DateTime SentAt { get; set; } = DateTime.UtcNow;
        public DateTime? EditedAt { get; set; }

        public List<MessageAttachment> Attachments { get; set; } = new List<MessageAttachment>();
        public List<Message> Replies { get; set; } = new List<Message>();

        public IntraUserProfile? Sender => !string.IsNullOrEmpty(SenderId) ? new IntraUserProfile { UserName = SenderId, FullUserName = SenderName ?? SenderId } : null;
    }

    public class MessageAttachment
    {
        public int Id { get; set; }
        public long MessageId { get; set; }
        public string FileName { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public long FileSize { get; set; }
        public string? ContentType { get; set; }
        public DateTime UploadedAt { get; set; } = DateTime.UtcNow;
    }

    public class DirectChatSummary
    {
        public string OtherUserId { get; set; } = string.Empty;
        public string OtherUserName { get; set; } = string.Empty;
        public string? OtherUserFullName { get; set; }
        public string? Designation { get; set; }
        public string? DepartmentName { get; set; }
        public UserStatus Status { get; set; } = UserStatus.Offline;
        public string? LastMessage { get; set; }
        public DateTime LastMessageTime { get; set; }
        public DateTime LastMessageAt { get => LastMessageTime; set => LastMessageTime = value; }
        public int UnreadCount { get; set; }
    }
}
